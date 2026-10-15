from docx import Document

d = Document(r"C:\Users\Admin\Desktop\New folder\SD_26_GIMASYS_01_TÀI_LIỆU_MÔ_TẢ_DỮ_LIỆU_DỰ_ÁN_GMS.docx")
for name in ["Normal", "Heading 1", "Heading 2", "Title"]:
    s = d.styles[name]
    f = s.font
    p = s.paragraph_format
    print(name, f.name, f.size.pt if f.size else None, f.bold,
          str(f.color.rgb) if f.color and f.color.rgb else None,
          p.space_before.pt if p.space_before else None,
          p.space_after.pt if p.space_after else None,
          p.line_spacing)
for idx in [0, 6, 7, 30, 31, 32, 41, 42]:
    p = d.paragraphs[idx]
    print("P", idx, p.text[:50], [(r.font.name, r.font.size.pt if r.font.size else None, r.bold, str(r.font.color.rgb) if r.font.color and r.font.color.rgb else None) for r in p.runs])
