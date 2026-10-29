import win32com.client

word = win32com.client.Dispatch('Word.Application')
word.Visible = False
doc_path = r'C:\Users\ganesh jadhav\Downloads\CAPSTONE\AI-Student-Agent\AI_Student_Agent_Black_Book_Final.docx'
wdoc = word.Documents.Open(doc_path)

# Configure Section 2 (Roman)
s2 = wdoc.Sections(2)
s2.Headers(1).LinkToPrevious = False
s2.Footers(1).LinkToPrevious = False
s2.Footers(1).PageNumbers.RestartNumberingAtSection = True
s2.Footers(1).PageNumbers.StartingNumber = 1
s2.Footers(1).PageNumbers.NumberStyle = 2 # Lowercase Roman
s2.Footers(1).PageNumbers.Add(2) # Right align

# Configure Section 3 (Arabic starting at 1)
s3 = wdoc.Sections(3)
s3.Headers(1).LinkToPrevious = False
s3.Headers(1).Range.Text = "AI Student Agent: An Autonomous AI Learning Assistant for Real-Time Online Classes\t\tAcademic Year 2026-2027"
s3.Headers(1).Range.Font.Name = 'Times New Roman'
s3.Headers(1).Range.Font.Size = 10
s3.Footers(1).LinkToPrevious = False
s3.Footers(1).PageNumbers.RestartNumberingAtSection = True
s3.Footers(1).PageNumbers.StartingNumber = 1
s3.Footers(1).PageNumbers.NumberStyle = 0 # Arabic
s3.Footers(1).PageNumbers.Add(2) # Right align

wdoc.Repaginate()

print("--- SECTION 2 PRELIMINARY PAGES ---")
for p in s2.Range.Paragraphs:
    txt = p.Range.Text.strip()
    if p.Range.Tables.Count == 0:
        for t in ["ACKNOWLEDGEMENT", "ABSTRACT", "TABLE OF CONTENTS", "LIST OF FIGURES", "LIST OF TABLES", "ABBREVIATIONS"]:
            if txt == t:
                # Convert adjusted page number to Roman string
                pg = p.Range.Information(1)
                print(f"  {t:20}: Page {pg}")

print("\n--- SECTION 3 CHAPTERS AND SECTIONS ---")
chap_targets = [
    ("Chapter 1", "Chapter 1\rIntroduction"),
    ("  1.1", "1.1 Background of the Project Topic"),
    ("  1.2", "1.2 Motivation and Scope of the Report"),
    ("  1.3", "1.3 Problem Statement"),
    ("  1.4", "1.4 Salient Contributions"),
    ("  1.5", "1.5 Organization of the Report"),
    ("Chapter 2", "Chapter 2\rLiterature Survey"),
    ("  2.1", "2.1 Introduction to the Literature Survey"),
    ("  2.2", "2.2 Detailed Review of Existing Research"),
    ("Chapter 3", "Chapter 3\rMethodology and Implementation"),
    ("  3.1", "3.1 Block Diagram and Architectural Framework"),
    ("  3.2", "3.2 Hardware Description"),
    ("  3.3", "3.3 Software Description, Flowchart and Algorithms"),
    ("  3.4", "3.4 Technical Implementation Challenges and Solutions"),
    ("Chapter 4", "Chapter 4\rResults and Analysis"),
    ("  4.1", "4.1 System Verification and Functional Test Matrix"),
    ("  4.2", "4.2 Qualitative Demonstration and Live Session Outcomes"),
    ("  4.3", "4.3 Latency, Resource Consumption, and Performance Analysis"),
    ("  4.4", "4.4 Comparative Evaluation of Architecture Trade-offs"),
    ("  4.5", "4.5 Project Development Timeline and Milestone Tracking"),
    ("Chapter 5", "Chapter 5\rAdvantages, Limitations and Applications"),
    ("  5.1", "5.1 Advantages"),
    ("  5.2", "5.2 Limitations"),
    ("  5.3", "5.3 Applications"),
    ("Chapter 6", "Chapter 6\rConclusion and Future Scope"),
    ("  6.1", "6.1 Conclusion"),
    ("  6.2", "6.2 Future Scope"),
    ("References", "References"),
    ("Appendix I", "Appendix I: Sample Code Excerpts"),
    ("Appendix II", "Appendix II: Technical Specifications and Component Data Sheets"),
    ("Appendix III", "Appendix III: Comprehensive List of Components"),
    ("Appendix IV", "Appendix IV: List of Papers Presented and Published"),
]

for label, search_str in chap_targets:
    found = False
    for p in s3.Range.Paragraphs:
        if p.Range.Tables.Count > 0:
            continue
        txt = p.Range.Text.strip()
        if search_str.lower() in txt.lower():
            print(f"  {label:20}: Page {p.Range.Information(1)}")
            found = True
            break
    if not found:
        print(f"  {label:20}: NOT FOUND")

wdoc.Close(False)
word.Quit()
