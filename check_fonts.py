from pptx import Presentation

prs = Presentation(r'C:\Users\ganesh jadhav\Downloads\MedIntel_AI_Final_Presentation.pptx')

# Check slide master and layouts
slide = prs.slides[0]
print("Slide 0 layout name:", slide.slide_layout.name)
print("Slide master layouts:", [l.name for l in prs.slide_layouts])

# Check font on Title in slide 1
for p in prs.slides[1].shapes[0].text_frame.paragraphs:
    print(f"Title font: name={p.font.name}, size={p.font.size}, color={p.font.color}")
    for r in p.runs:
        print(f"  Run: text='{r.text}', font={r.font.name}, size={r.font.size}, color={r.font.color.rgb if r.font.color and hasattr(r.font.color, 'rgb') else 'None'}")

# Check text in slide 2 body
for p in prs.slides[1].shapes[2].text_frame.paragraphs:
    print(f"Body p: font={p.font.name}, size={p.font.size}")
    for r in p.runs:
        print(f"  Run: text='{r.text[:30]}', font={r.font.name}, size={r.font.size}, bold={r.font.bold}, color={r.font.color.rgb if r.font.color and hasattr(r.font.color, 'rgb') else 'None'}")
