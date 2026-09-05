"""PDF 上传解析：pdfplumber 逐页提取纯文本（设计参考 yuyuanweb/ai-ppt 的 backend/app/ingest/pdf.py）。

设计理由——三类失败路径（加密 / 扫描件 / 打不开）全部明确抛 ValueError、绝不静默返回空文本：
下游大纲生成以这份文本为唯一素材，解析失败若返回空串，流水线会「看似成功」地
生成一份空洞 PPT，用户要等整条链路跑完才发现；尽早失败，用户才能立刻换文件重传。
"""

import io

import pdfplumber
from pdfminer.pdfdocument import PDFEncryptionError


def _cause(exc: Exception) -> BaseException:
    """取 pdfminer 原始异常。pdfplumber 统一用 raise PdfminerException(原异常) 包装，args[0] 即原异常。"""
    arg = exc.args[0] if exc.args else None
    return arg if isinstance(arg, BaseException) else exc


def parse_pdf(data: bytes) -> str:
    """把 PDF 字节流解析成纯文本，页与页之间以换行符连接。

    - 加密件：先尝试空密码（多数加密 PDF 只是权限限制、并无打开密码），
      解不开则抛 ValueError("PDF 已加密，无法解析")
    - 扫描件：所有页都提取不到文本时抛 ValueError（无文字层的图片型 PDF）
    - 非 PDF / 损坏文件：抛 ValueError(f"PDF 无法打开: {原因}")
    """
    try:
        # password="" 即「先尝试空密码」：pdfplumber 0.11.10 依赖的 pdfminer.six
        # 20260107 已移除 PDF.is_encrypted / PDF.decrypt 旧入口，空密码的尝试
        # 在 open 阶段一次完成——解开则正常返回，解不开抛 PDFPasswordIncorrect
        # （被包装进 PdfminerException.args[0]）。
        pdf = pdfplumber.open(io.BytesIO(data), password="")
    except Exception as e:
        if isinstance(_cause(e), PDFEncryptionError):
            raise ValueError("PDF 已加密，无法解析") from e
        raise ValueError(f"PDF 无法打开: {e}") from e

    with pdf:
        # extract_text 对空白页 / 纯图片页返回空串或 None；空页不参与 join，
        # 避免页与页之间留下多余空行
        texts: list[str] = []
        for page in pdf.pages:
            text = (page.extract_text() or "").strip()
            if text:
                texts.append(text)
        if not texts:
            raise ValueError("PDF 中没有可提取的文本（可能是扫描件）")
        return "\n".join(texts)
