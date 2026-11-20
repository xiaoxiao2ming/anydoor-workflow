# -*- coding: utf-8 -*-
"""PDF → Word 转换脚本（client-base PDF转word 节点专用）

引擎 = 百度开源 PaddleOCR（飞桨）PP-StructureV3 文档解析管线：
版面分析 + OCR + 表格识别 + 版面恢复，产出可编辑 docx。

用法：
    python run_p2w.py input.pdf output.docx [--lang ch|en]
    python run_p2w.py --serve   （长驻模式：行式 JSON 协议，管线只加载一次）

依赖（本机一次性安装，套件中心「装入套件目录」自动完成）：
    pip install paddlepaddle==3.2.1 "paddleocr[doc-parser]" python-docx docxcompose
    ★ paddlepaddle 锁 3.2.x：3.3 存在 OneDNN/PIR 推理回归
    ★ python-docx/docxcompose：save_to_word 落盘与多页合并必需
性能：默认关闭 公式识别/文档方向/文本行方向/弯曲矫正 四个辅助模型
    （燃气/工程文档用不上公式，关掉后加载快 10s+、每页推理快 30-50%、内存占用大降——
    PP-FormulaNet_plus-L 单模型 728M 是内存大户）。需要公式识别时改 USE_FORMULA=True。
模型：首次运行自动下载（1-2GB，缓存 ~/.paddlex，之后离线可用）。
输出约定：成功打印 P2W_OK，失败打印 P2W_ERROR: 原因 并以非零码退出。
"""
USE_FORMULA = False
import argparse
import glob
import os
import shutil
import sys
import tempfile

# ★ stdout/stderr 统一 UTF-8：Windows 默认 cp936，输出中文（路径/错误原因）会 UnicodeEncodeError，
#   而 Electron 侧按 UTF-8 解码读回，编码一致才不乱码
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


def save_docx(res_list, output: str, work_dir: str) -> int:
    """逐页 save_to_word 落盘 → 收集 → 合并为单文件（0=成功，5=失败）。
    ★ 实测行为：save_path 当目录用（写 <dir>/<输入名>_N.docx），目录必须先建好。"""
    import re
    base = os.path.join(work_dir, 'page')
    os.makedirs(base, exist_ok=True)
    for res in res_list:
        try:
            res.save_to_word(base)
        except TypeError:
            res.save_to_word(save_path=base)
    parts = glob.glob(os.path.join(work_dir, '**', '*.docx'), recursive=True)

    def _idx(f: str) -> int:
        m = re.search(r'_(\d+)\.docx$', f)
        return int(m.group(1)) if m else 0

    parts.sort(key=_idx)
    if not parts:
        listing = []
        for root, _dirs, files in os.walk(work_dir):
            listing += [os.path.join(root, f) for f in files]
        raise RuntimeError('引擎未落盘任何 Word——work=' + repr(listing[:8]) + ' cwd=' + os.getcwd() + ' cwd-docx=' + repr(glob.glob('*.docx')[:5]))
    if len(parts) == 1:
        shutil.move(parts[0], output)
    else:
        from docx import Document
        from docxcompose.composer import Composer
        composer = Composer(Document(parts[0]))
        for f in parts[1:]:
            composer.append(Document(f))
        composer.save(output)
    if os.path.isfile(output) and os.path.getsize(output) > 0:
        print('P2W_OK ' + output + '（合并 ' + str(len(parts)) + ' 页）')
        return 0
    print('P2W_ERROR: 产物未生成', file=sys.stderr)
    return 5


def serve() -> int:
    """★ 长驻引擎模式（--serve）：行式 JSON 协议—— stdin 每行一条
    {input, output, lang, profile}，stdout 回一行 {ok, err?, pages?}。
    管线只加载一次（首请求时按 lang+profile 惰性创建；任一变了自动重建），
    免去每次转换 30-60s 的 import+模型加载。stderr 留给引擎日志。"""
    import json
    global work
    work = None  # ★ 必须先初始化：首请求在 mkdtemp 之前失败时，finally 里 rmtree(None) 会 NameError，
    #   异常从 finally 抛出 = 整个长驻进程崩溃退出（用户侧只看到「python 退出码 -1」，真实原因全丢）
    engine = None
    engine_key = None
    print('P2W_READY', flush=True)
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
            inp, out = req['input'], req['output']
            lang = req.get('lang', 'ch')
            profile = req.get('profile', 'pro')
            if profile not in PROFILES:
                profile = 'pro'
            if not os.path.isfile(inp):
                raise RuntimeError('输入文件不存在: ' + inp)
            key = lang + '|' + profile
            if engine is None or engine_key != key:
                engine = make_engine(lang, profile)
                engine_key = key
            work = tempfile.mkdtemp(prefix='p2w-save-')
            res_list = list(engine.predict(inp))
            if not res_list:
                raise RuntimeError('PDF 无可识别内容')
            rc = save_docx(res_list, out, work)
            if rc == 0:
                print(json.dumps({'ok': True, 'pages': len(res_list)}), flush=True)
            else:
                print(json.dumps({'ok': False, 'err': '产物未生成'}), flush=True)
        except Exception as e:
            print(json.dumps({'ok': False, 'err': ('%s: %s' % (type(e).__name__, e)).strip(': ')}, ensure_ascii=False), flush=True)
        finally:
            if work:
                shutil.rmtree(work, ignore_errors=True)
                work = None
    return 0


# ── 速度档位（2026-10-26 用户问「有没有容差设置」）──────────────────
# pro 精准（缺省）= server OCR + 表格结构还原（9 模型全链，~13s/页）
# std 均衡 = server OCR + 关表格还原（省 4 个表格模型，表格退化为纯文本行）
# fast 极速 = mobile OCR + 关表格还原（模型最小，估 2-5s/页，长文档首选）
PROFILES = ('fast', 'std', 'pro')


def make_engine(lang: str, profile: str = 'pro'):
    """创建管线（辅助模型默认关闭——速度优先，见文件头说明）"""
    from paddleocr import PPStructureV3
    kw = dict(
        lang=lang,
        use_formula_recognition=USE_FORMULA,
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
    )
    if profile == 'fast':
        kw['use_table_recognition'] = False
        kw['text_detection_model_name'] = 'PP-OCRv5_mobile_det'
        kw['text_recognition_model_name'] = 'PP-OCRv5_mobile_rec'
    elif profile == 'std':
        kw['use_table_recognition'] = False
    return PPStructureV3(**kw)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument('input', help='输入 PDF 路径')
    ap.add_argument('output', help='输出 docx 路径')
    ap.add_argument('--lang', default='ch', choices=['ch', 'en'], help='识别语言（缺省 ch）')
    ap.add_argument('--profile', default='pro', choices=list(PROFILES), help='速度档位（缺省 pro 精准）')
    a = ap.parse_args()

    if not os.path.isfile(a.input):
        print('P2W_ERROR: 输入文件不存在: ' + a.input, file=sys.stderr)
        return 2

    try:
        from paddleocr import PPStructureV3
    except Exception as e:
        print('P2W_ERROR: 未安装 PaddleOCR（pip install paddlepaddle==3.2.1 "paddleocr[doc-parser]"）: ' + str(e), file=sys.stderr)
        return 3

    out_dir = os.path.dirname(os.path.abspath(a.output)) or '.'
    os.makedirs(out_dir, exist_ok=True)
    work = tempfile.mkdtemp(prefix='p2w-save-')

    try:
        engine = make_engine(a.lang, a.profile)
        res_list = list(engine.predict(a.input))
        if not res_list:
            print('P2W_ERROR: 引擎未产出任何结果（PDF 无可识别内容？）', file=sys.stderr)
            return 4

        # ① 管线级整档合并（PaddleX 若提供则优先——零页序风险）
        try:
            engine.save_to_merging_word(res_list, a.output)
            if os.path.isfile(a.output):
                print('P2W_OK ' + a.output)
                return 0
        except AttributeError:
            pass

        # ② 逐页 save_to_word——实测：save_path 当目录用（写 <dir>/<输入名>_N.docx）
        #    且不会自动建目录（目录缺失时静默不出文件）→ 先 makedirs
        base = os.path.join(work, 'page')
        os.makedirs(base, exist_ok=True)
        for res in res_list:
            try:
                res.save_to_word(base)
            except TypeError:
                res.save_to_word(save_path=base)
        import re
        parts = glob.glob(os.path.join(work, '**', '*.docx'), recursive=True)

        def _idx(f: str) -> int:
            m = re.search(r'_(\d+)\.docx$', f)
            return int(m.group(1)) if m else 0

        parts.sort(key=_idx)
        if not parts:
            print('P2W_ERROR: 引擎未落盘任何 Word（save_to_word 行为未知）', file=sys.stderr)
            return 5
        if len(parts) == 1:
            shutil.move(parts[0], a.output)
        else:
            from docx import Document
            from docxcompose.composer import Composer
            composer = Composer(Document(parts[0]))
            for f in parts[1:]:
                composer.append(Document(f))
            composer.save(a.output)
        if os.path.isfile(a.output) and os.path.getsize(a.output) > 0:
            print('P2W_OK ' + a.output + '（合并 ' + str(len(parts)) + ' 页）')
            return 0
        print('P2W_ERROR: 产物未生成', file=sys.stderr)
        return 5
    except Exception as e:
        print('P2W_ERROR: ' + str(e), file=sys.stderr)
        return 1
    finally:
        shutil.rmtree(work, ignore_errors=True)


if __name__ == '__main__':
    if '--serve' in sys.argv:
        sys.exit(serve())
    sys.exit(main())
