from pptx import Presentation

prs = Presentation(r'C:\Users\ganesh jadhav\Downloads\MedIntel_AI_Final_Presentation.pptx')

print("Slide master background:")
try:
    bg = prs.slide_master.background
    print("Master bg type:", bg.fill.type)
except Exception as e:
    print("Master bg err:", e)

# Check each slide's background
for i in [0, 1, 5, 10]:
    s = prs.slides[i]
    try:
        print(f"Slide {i+1} bg type: {s.background.fill.type}")
    except Exception as e:
        print(f"Slide {i+1} bg err: {e}")
