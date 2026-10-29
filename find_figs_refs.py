import pypdf

reader = pypdf.PdfReader('test_sec3_all.pdf')

figs = [
    'Figure 3.1', 'Figure 3.2', 'Figure 3.3', 'Figure 3.4', 'Figure 3.5',
    'Figure 3.6', 'Figure 3.7', 'Figure 4.1', 'Figure 4.2', 'Figure 4.3'
]

print("--- FIGURES ---")
for f in figs:
    for idx in range(15, len(reader.pages)):
        txt = reader.pages[idx].extract_text()
        if f in txt:
            sec3_pg = idx - 15 + 1
            print(f"{f}: Sec3 Page {sec3_pg} (Abs Page {idx+1})")
            break

print("\n--- REFERENCES ---")
for idx in range(65, len(reader.pages)):
    txt = reader.pages[idx].extract_text()
    if "REFERENCES" in txt or "References" in txt:
        sec3_pg = idx - 15 + 1
        print(f"References Heading: Sec3 Page {sec3_pg} (Abs Page {idx+1})")
        break
