import win32com.client
import os
import time

ppt_path = r"C:\Users\87090\Doubao\chats\2026-08-18\new-chat-1\项目1-Gymshark品牌研究\report\Gymshark品牌深度研究报告.pptx"
pdf_path = r"C:\Users\87090\Doubao\chats\2026-08-18\new-chat-1\项目1-Gymshark品牌研究\report\Gymshark品牌深度研究报告.pdf"

print(f"正在打开PPT: {ppt_path}")
powerpoint = win32com.client.Dispatch("PowerPoint.Application")
powerpoint.Visible = 1

try:
    presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
    print(f"PPT已打开，共 {len(presentation.Slides)} 页")
    
    # 32 = ppSaveAsPDF
    presentation.SaveAs(pdf_path, 32)
    print(f"PDF已生成: {pdf_path}")
    
    presentation.Close()
    print("PPT已关闭")
except Exception as e:
    print(f"错误: {e}")
finally:
    powerpoint.Quit()
    print("PowerPoint已退出")

# 验证文件
if os.path.exists(pdf_path):
    size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
    print(f"PDF文件大小: {size_mb:.2f} MB")
else:
    print("PDF生成失败！")
