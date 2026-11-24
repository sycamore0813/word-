# -*- coding: utf-8 -*-
"""调用本机 Word（COM）更新目录/页码域，并可导出 PDF 用于检查效果。

不依赖 pywin32：改用 PowerShell 调 COM，脚本本身是 ASCII，路径通过环境变量传入，
避免中文路径在命令行里被转码搞坏。
"""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from typing import Optional, Tuple

_PS = r"""
$ErrorActionPreference = 'Stop'
$docx = $env:WF_DOCX
$pdf  = $env:WF_PDF
$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    try { $word.DisplayAlerts = 0 } catch {}
    try { $word.Options.UpdateFieldsAtPrint = $true } catch {}
    $doc = $word.Documents.Open($docx, $false, $false, $false)
    try { $doc.Fields.Update() | Out-Null } catch { Write-Output ("WARN fields: " + $_.Exception.Message) }
    try {
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() | Out-Null }
    } catch { Write-Output ("WARN toc: " + $_.Exception.Message) }
    try { $doc.Repaginate() } catch {}
    $doc.Save()
    if ($pdf -and $pdf.Length -gt 0) {
        try {
            $doc.ExportAsFixedFormat($pdf, 17) | Out-Null
            Write-Output ("PDF_OK " + $pdf)
        } catch { Write-Output ("WARN pdf: " + $_.Exception.Message) }
    }
    Write-Output "WORD_OK"
}
catch {
    Write-Output ("WORD_FAIL " + $_.Exception.Message)
}
finally {
    if ($doc -ne $null) { try { $doc.Close(0) } catch {} }
    if ($word -ne $null) { try { $word.Quit() } catch {} }
}
"""


def _powershell_exe() -> str:
    for name in ("pwsh.exe", "powershell.exe"):
        path = shutil.which(name)
        if path:
            return path
    return "powershell.exe"


def update_fields(docx_path: str, pdf_path: Optional[str] = None,
                  timeout: int = 300) -> Tuple[bool, str]:
    """用 Word 更新域并保存；pdf_path 非空时同时导出 PDF。"""
    docx_path = os.path.abspath(docx_path)
    if not os.path.exists(docx_path):
        return False, "文件不存在：%s" % docx_path
    script = os.path.join(tempfile.gettempdir(), "wordfmt_update_fields.ps1")
    try:
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(_PS)
    except Exception as e:
        return False, "写入临时脚本失败：%s" % e
    env = dict(os.environ)
    env["WF_DOCX"] = docx_path
    env["WF_PDF"] = os.path.abspath(pdf_path) if pdf_path else ""
    try:
        proc = subprocess.run([_powershell_exe(), "-NoProfile", "-ExecutionPolicy", "Bypass",
                               "-File", script],
                              capture_output=True, timeout=timeout, env=env,
                              text=True, encoding="utf-8", errors="replace")
    except subprocess.TimeoutExpired:
        return False, "Word 处理超时（%d 秒），可能弹了对话框" % timeout
    except Exception as e:
        return False, "调用 Word 失败：%s" % e
    out = (proc.stdout or "") + (proc.stderr or "")
    if "WORD_OK" in out:
        return True, out.strip()
    return False, (out.strip() or "未知错误")


def word_installed() -> bool:
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe"):
            return True
    except Exception:
        return False
