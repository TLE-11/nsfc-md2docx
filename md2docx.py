#!/usr/bin/env python3
"""Markdown -> Word（中文本子排版）一键流水线。跨平台，Windows / macOS / Linux 通用。

  python md2docx.py 输入.md [-o 输出.docx] [选项]

只依赖 Python 3 标准库 + pandoc 可执行文件。Windows 上不需要 bash / WSL。

主要选项：
  --math-mode omml|latex   公式形态。omml=Word 原生公式（默认，不需插件）；
                           latex=保留 LaTeX 源码，交给 MathType/AxMath 批量转换
  --number all|tag|none    行间公式右编号范围。默认 all（所有行间公式都编号）
  --number-style plain|chapter   (1)(2)... 还是 (2-1)(3-1)...
  --fonts windows|macos    样式模板字体档位。windows=宋体/黑体（出稿用）；
                           macos=Songti SC/Heiti SC（本机预览用，macOS 没有宋体）
  --assets 目录 ...        图片附件搜索目录（Obsidian 附件目录）
  --pdf                    额外用 LibreOffice 渲染一份 PDF 便于肉眼检查
  --keep-temp              保留中间文件，便于排查
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))

# pandoc 扩展说明：
#   +tex_math_dollars       $...$ / $$...$$ 识别为公式
#   -yaml_metadata_block    否则文中的 --- 被误判成 YAML 头，直接报错退出
#   -simple_tables
#   -multiline_tables       否则 --- 被当成表格头分隔符，把后面整段正文吞成表格纯文本
PANDOC_FROM = ('markdown+tex_math_dollars-yaml_metadata_block'
               '-simple_tables-multiline_tables')


def which_pandoc():
    p = shutil.which('pandoc')
    if p:
        return p
    # Windows 常见安装位置（装完没重启终端时 PATH 里可能还没有）
    for c in (os.path.expandvars(r'%LOCALAPPDATA%\Pandoc\pandoc.exe'),
              r'C:\Program Files\Pandoc\pandoc.exe'):
        if os.path.exists(c):
            return c
    sys.exit('找不到 pandoc。装法：\n'
             '  Windows : winget install --id JohnMacFarlane.Pandoc\n'
             '            或 choco install pandoc\n'
             '  macOS   : brew install pandoc\n'
             '  Linux   : apt install pandoc')


def which_soffice():
    for c in ('soffice', 'libreoffice'):
        p = shutil.which(c)
        if p:
            return p
    for c in ('/Applications/LibreOffice.app/Contents/MacOS/soffice',
              r'C:\Program Files\LibreOffice\program\soffice.exe'):
        if os.path.exists(c):
            return c
    return None


def run(cmd, label):
    r = subprocess.run(cmd)
    if r.returncode != 0:
        sys.exit('[pipe] %s 失败（退出码 %d）' % (label, r.returncode))


def main():
    ap = argparse.ArgumentParser(add_help=True)
    ap.add_argument('src')
    ap.add_argument('-o', '--out')
    ap.add_argument('--math-mode', choices=['omml', 'latex'], default='omml')
    ap.add_argument('--number', choices=['all', 'tag', 'none'], default='all')
    ap.add_argument('--number-style', choices=['plain', 'chapter'], default='plain')
    ap.add_argument('--fonts', choices=['windows', 'macos'], default='windows')
    ap.add_argument('--assets', nargs='*', default=[])
    ap.add_argument('--reference', help='自定义 reference.docx，不给就自动生成')
    ap.add_argument('--pdf', action='store_true')
    ap.add_argument('--keep-temp', action='store_true')
    a = ap.parse_args()

    if not os.path.exists(a.src):
        sys.exit('找不到输入文件: %s' % a.src)
    out = a.out or (os.path.splitext(a.src)[0] + '.docx')
    pandoc = which_pandoc()
    py = sys.executable

    # ---- 样式模板 ----
    ref = a.reference
    if not ref:
        ref = os.path.join(HERE, 'reference-%s.docx' % a.fonts)
        if not os.path.exists(ref):
            print('[pipe] 生成样式模板（字体档位: %s）' % a.fonts)
            run([py, os.path.join(HERE, 'make_reference.py'), ref,
                 '--font-profile', a.fonts], 'make_reference')

    tmp = tempfile.mkdtemp(prefix='md2docx_')
    try:
        pre_md = os.path.join(tmp, 'pre.md')
        math_json = os.path.join(tmp, 'math.json')
        step1 = os.path.join(tmp, 'step1.docx')

        print('[pipe] 1/4 预处理（公式 %s，编号 %s/%s）'
              % (a.math_mode, a.number, a.number_style))
        cmd = [py, os.path.join(HERE, 'pre.py'), a.src, pre_md,
               '--math-mode', a.math_mode, '--math-json', math_json,
               '--number', a.number, '--number-style', a.number_style]
        if a.assets:
            cmd += ['--assets'] + a.assets
        run(cmd, 'pre.py')

        print('[pipe] 2/4 pandoc 转换')
        rp = [os.path.dirname(os.path.abspath(a.src))] + list(a.assets)
        run([pandoc, pre_md, '--from', PANDOC_FROM, '--to', 'docx',
             '--reference-doc', ref, '--resource-path', os.pathsep.join(rp),
             '-o', step1], 'pandoc')
        if not os.path.exists(step1):
            sys.exit('[pipe] pandoc 没有产出文件')

        print('[pipe] 3/4 后处理')
        run([py, os.path.join(HERE, 'post.py'), step1, out,
             '--math-json', math_json, '--number-style', a.number_style], 'post.py')

        print('[pipe] 4/4 校验')
        r = subprocess.run([py, os.path.join(HERE, 'verify.py'), a.src, out,
                            '--math-mode', a.math_mode, '--number', a.number])
        if r.returncode != 0:
            sys.exit('[pipe] 校验未通过，产物可能有问题，请看上面的失败项')
    finally:
        if a.keep_temp:
            print('[pipe] 中间文件保留在 %s' % tmp)
        else:
            shutil.rmtree(tmp, ignore_errors=True)

    print('[pipe] 完成 -> %s' % out)

    if a.pdf:
        so = which_soffice()
        if not so:
            print('[pipe] 未找到 LibreOffice，跳过 PDF 预览')
        else:
            d = os.path.dirname(os.path.abspath(out)) or '.'
            subprocess.run([so, '--headless', '--convert-to', 'pdf', out,
                            '--outdir', d], stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL)
            print('[pipe] PDF 预览 -> %s' % (os.path.splitext(out)[0] + '.pdf'))

    if a.math_mode == 'latex':
        print("""
[pipe] 公式现在是蓝色等宽的 LaTeX 源码，需在 Windows + MS Word 上做最后一步：
       1. 打开 docx，按 Alt+F11 进 VBA 编辑器
       2. 文件 -> 导入文件，选 mathtype_axmath.bas（仓库根目录）
       3. 运行 ConvertLatexToMathType 或 ConvertLatexToAxMath
       详见 README_cn.md 的「支持 MathType / AxMath」一节""")


if __name__ == '__main__':
    main()
