#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ensure_fonts.py — 检测本机是否缺少公文字体，缺失则自动下载并安装（用户级，无需管理员）。

用法:
    python ensure_fonts.py                     # 检测 + 安装所有缺失的公文字体
    python ensure_fonts.py --check             # 只检测，报告缺失情况
    python ensure_fonts.py --font 黑体         # 只安装指定字体

适用: Windows + MS Office COM 场景。字体缺失时 Word 会用相似字体替换显示（见 SKILL.md
"Font substitution"），为保证公文观感一致，生成/修改文档前先运行本脚本。

下载源（GitHub 公开仓库，仅用于公文字体学习/办公场景）:
    - guorenxi/MacFonts     (仿宋_GB2312, 楷体_GB2312, 方正小标宋简体, 方正小标宋_GBK)
    - Mackerly/fonts        (仿宋_GB2312, 方正仿宋, 方正小标宋_GBK, 方正楷体简体, 黑体)
    注意: 国内访问 GitHub 常被屏蔽/限速。本脚本对 GitHub 源采用"快速失败"策略:
    连接超时 10s / 总下载限时 30s（可用 FONT_DL_CONNECT_TIMEOUT / FONT_DL_TOTAL_TIMEOUT 环境变量
    覆盖），超时即跳过该源不傻等；全部失败时提示设置代理后重试。

安装方式: 复制到 %LOCALAPPDATA%\\Microsoft\\Windows\\Fonts + HKCU 注册表 + AddFontResource +
广播 WM_FONTCHANGE。用户级安装，无需管理员权限；Word/Excel/PowerPoint 重启后生效。
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import winreg
import ctypes
from ctypes import wintypes

# ---------- 字体清单：系统字体名 -> 候选下载源 ----------
# key 用注册表字体名（含 "(TrueType)" 匹配用短名）
FONT_SOURCES = {
    "仿宋_GB2312": [
        "https://raw.githubusercontent.com/guorenxi/MacFonts/master/fonts/MacFSGB2312.ttf",
        "https://raw.githubusercontent.com/Mackerly/fonts/main/%E4%BB%BF%E5%AE%8B_GB2312.TTF",
    ],
    "楷体_GB2312": [
        "https://raw.githubusercontent.com/guorenxi/MacFonts/master/fonts/MacKTGB2312.ttf",
    ],
    "方正小标宋简体": [
        "https://raw.githubusercontent.com/guorenxi/MacFonts/master/fonts/MacFZXBSJT.ttf",
        "https://raw.githubusercontent.com/Mackerly/fonts/main/%E6%96%B9%E6%AD%A3%E5%B0%8F%E6%A0%87%E5%AE%8B%E7%AE%80.TTF",
    ],
    "方正小标宋_GBK": [
        "https://raw.githubusercontent.com/guorenxi/MacFonts/master/fonts/MacFZXBSGBK.ttf",
        "https://raw.githubusercontent.com/Mackerly/fonts/main/%E6%96%B9%E6%AD%A3%E5%B0%8F%E6%A0%87%E5%AE%8B_GBK.TTF",
    ],
    "方正楷体_GBK": [
        "https://raw.githubusercontent.com/Mackerly/fonts/main/%E6%96%B9%E6%AD%A3%E6%A5%B7%E4%BD%93%E7%AE%80%E4%BD%93.TTF",
    ],
    "黑体": [
        "https://raw.githubusercontent.com/Mackerly/fonts/main/%E6%96%B9%E6%AD%A3%E7%B2%97%E9%BB%91%E5%AE%8B%E7%AE%80%E4%BD%93.ttf",
    ],
    "方正仿宋_GB2312": [
        "https://raw.githubusercontent.com/Mackerly/fonts/main/%E6%96%B9%E6%AD%A3%E4%BB%BF%E5%AE%8B_GB2312.ttf",
        "https://raw.githubusercontent.com/Mackerly/fonts/main/%E6%96%B9%E6%AD%A3%E4%BB%BF%E5%AE%8B%E7%AE%80%E4%BD%93.TTF",
    ],
}


def list_installed_fonts():
    """枚举本机已安装字体名（HKLM + HKCU 注册表，无需管理员）"""
    fonts = set()
    for hive, path in [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
    ]:
        try:
            k = winreg.OpenKey(hive, path)
            i = 0
            while True:
                try:
                    name, _, _ = winreg.EnumValue(k, i)
                    fonts.add(name)
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(k)
        except OSError:
            pass
    return fonts


def is_installed(font_name, fonts):
    """按字体名模糊匹配（注册表名常带 '(TrueType)' 后缀）"""
    fn = font_name.lower()
    return any(fn in f.lower() for f in fonts)


# 下载保护参数（国内 GitHub 常被屏蔽/限速，超时即放弃该源，不傻等）
# 可通过环境变量覆盖：FONT_DL_CONNECT_TIMEOUT（连接超时秒，默认 10）、FONT_DL_TOTAL_TIMEOUT（总下载限时秒，默认 30）
CONNECT_TIMEOUT = float(os.environ.get("FONT_DL_CONNECT_TIMEOUT", "10"))
TOTAL_TIMEOUT = float(os.environ.get("FONT_DL_TOTAL_TIMEOUT", "30"))


def download(url, dest_dir):
    """下载字体文件，返回本地路径；失败/超时返回 None。

    如果来源是 GitHub（raw.githubusercontent.com / github.com）且连接超时或总下载
    超过限时，直接放弃该源（国内经常屏蔽 GitHub，快速失败比长时间挂起好）。
    """
    import time
    import urllib.parse

    is_github = "github.com" in url or "raw.githubusercontent.com" in url
    fname = os.path.basename(url.split("?")[0])
    if "%" in fname:  # URL 编码的文件名
        fname = urllib.parse.unquote(fname)
    dest = os.path.join(dest_dir, fname)
    start = time.time()
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=CONNECT_TIMEOUT) as r, open(dest, "wb") as f:
            while True:
                chunk = r.read(65536)
                if not chunk:
                    break
                f.write(chunk)
                # 总下载限时：超时即放弃（GitHub 源国内限速时特别有用）
                if time.time() - start > TOTAL_TIMEOUT:
                    raise TimeoutError(f"下载超过 {TOTAL_TIMEOUT:.0f}s 限时")
        if os.path.getsize(dest) < 10000:  # TTF 至少几十 KB，太小视为错误页
            os.remove(dest)
            return None
        return dest
    except Exception as e:
        try:
            if os.path.exists(dest):
                os.remove(dest)
        except OSError:
            pass
        tag = " [GitHub源超时/被墙，已跳过]" if is_github else ""
        print(f"  ! 下载失败 {url}: {e}{tag}")
        return None


def install_font_user(font_file):
    """用户级安装字体（无需管理员）：
    1. 复制到 %LOCALAPPDATA%\\Microsoft\\Windows\\Fonts
    2. 写 HKCU 注册表 Fonts 键
    3. AddFontResource 加载 + 广播 WM_FONTCHANGE
    """
    fonts_dir = os.path.join(os.environ["LOCALAPPDATA"], "Microsoft", "Windows", "Fonts")
    os.makedirs(fonts_dir, exist_ok=True)
    target = os.path.join(fonts_dir, os.path.basename(font_file))
    shutil.copy2(font_file, target)

    # 注册表：HKCU\\Software\\Microsoft\\Windows NT\\CurrentVersion\\Fonts
    key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Fonts"
    reg_name = os.path.splitext(os.path.basename(target))[0] + " (TrueType)"
    try:
        k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
        winreg.SetValueEx(k, reg_name, 0, winreg.REG_SZ, target)
        winreg.CloseKey(k)
    except OSError as e:
        print(f"  ! 注册表写入失败: {e}")

    # AddFontResource + WM_FONTCHANGE（立即生效）
    gdi32 = ctypes.windll.gdi32
    added = gdi32.AddFontResourceW(target)
    if added > 0:
        ctypes.windll.user32.SendMessageTimeoutW(
            wintypes.HWND(0xFFFF), 0x001D, 0, 0,
            0x0002, 2000, None)  # HWND_BROADCAST, WM_FONTCHANGE, SMTO_ABORTIFHUNG
    return target


def main():
    check_only = "--check" in sys.argv
    only_font = None
    if "--font" in sys.argv:
        i = sys.argv.index("--font")
        if i + 1 < len(sys.argv):
            only_font = sys.argv[i + 1]

    print("== 检测本机字体 ==")
    installed = list_installed_fonts()
    missing = []
    for name in FONT_SOURCES:
        if only_font and only_font not in name:
            continue
        if is_installed(name, installed):
            print(f"  ✅ {name}")
        else:
            print(f"  ❌ {name} 缺失")
            missing.append(name)

    if check_only:
        print(f"\n缺失 {len(missing)} 个字体" + ("，需要下载安装。" if missing else "，无需操作。"))
        return 0 if not missing else 1

    if not missing:
        print("\n所有公文字体已就绪。")
        return 0

    print(f"\n== 下载并安装缺失字体（共 {len(missing)} 个）==")
    workdir = tempfile.mkdtemp(prefix="fonts_dl_")
    installed_ok = []
    failed = []
    for name in missing:
        print(f"  {name}:")
        ok = False
        for url in FONT_SOURCES[name]:
            print(f"    尝试 {url.split('/')[-1][:40]}...")
            f = download(url, workdir)
            if f:
                try:
                    target = install_font_user(f)
                    print(f"    ✅ 已安装 → {target}")
                    installed_ok.append(name)
                    ok = True
                    break
                except Exception as e:
                    print(f"    ! 安装失败: {e}")
        if not ok:
            failed.append(name)

    shutil.rmtree(workdir, ignore_errors=True)
    print(f"\n完成: 安装成功 {len(installed_ok)} 个，失败 {len(failed)} 个。")
    if failed:
        print(f"失败: {failed}")
        print("提示: 这些字体可手动从 方正字库官网/单位字体库 获取后复制到 C:\\Windows\\Fonts。")
        print("提示: 若因 GitHub 被墙/限速导致下载失败，可设置代理环境变量后重试，")
        print("      如 set HTTPS_PROXY=http://127.0.0.1:10808 （或调整 FONT_DL_CONNECT_TIMEOUT/FONT_DL_TOTAL_TIMEOUT）。")
    print("注意: 已打开的 Word 需重启才能看到新字体。")
    return 0 if not failed else 1


if __name__ == "__main__":
    sys.exit(main())
