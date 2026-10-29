import os
import sys
import win32com.client

def main():
    doc_path = r"C:\Users\ganesh jadhav\Downloads\CAPSTONE\AI-Student-Agent\AI_Student_Agent_Black_Book_Final.docx"
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0

    try:
        wdoc = word.Documents.Open(doc_path)
        print(f"Total sections: {wdoc.Sections.Count}")

        # Section 1: Cover & Cert
        s1 = wdoc.Sections(1)
        s1.Headers(1).Range.Text = ""
        s1.Footers(1).Range.Text = ""

        # Section 2: Prelim (Roman)
        s2 = wdoc.Sections(2)
        s2.Headers(1).LinkToPrevious = False
        s2.Headers(1).Range.Text = ""
        s2.Footers(1).LinkToPrevious = False
        s2.Footers(1).PageNumbers.RestartNumberingAtSection = True
        s2.Footers(1).PageNumbers.StartingNumber = 1
        s2.Footers(1).PageNumbers.NumberStyle = 2  # Lowercase Roman
        s2.Footers(1).PageNumbers.Add(2)          # Right align

        # Section 3: Chapters (Arabic starting at 1)
        s3 = wdoc.Sections(3)
        s3.Headers(1).LinkToPrevious = False
        s3.Headers(1).Range.Text = "AI Student Agent: An Autonomous AI Learning Assistant for Real-Time Online Classes\t\tAcademic Year 2026-2027"
        s3.Headers(1).Range.Font.Name = "Times New Roman"
        s3.Headers(1).Range.Font.Size = 10
        s3.Footers(1).LinkToPrevious = False
        s3.Footers(1).PageNumbers.RestartNumberingAtSection = True
        s3.Footers(1).PageNumbers.StartingNumber = 1
        s3.Footers(1).PageNumbers.NumberStyle = 0  # Arabic
        s3.Footers(1).PageNumbers.Add(2)          # Right align

        wdoc.Repaginate()
        print(f"Repaginated. Total pages: {wdoc.ComputeStatistics(2)}")

        # Check pages in Section 2 using Find
        print("\n--- SECTION 2 PRELIMINARY PAGES ---")
        prelims = ["ACKNOWLEDGEMENT", "ABSTRACT", "TABLE OF CONTENTS", "LIST OF FIGURES", "LIST OF TABLES", "ABBREVIATIONS"]
        for p_name in prelims:
            rng = s2.Range
            find = rng.Find
            find.ClearFormatting()
            find.Text = p_name
            find.Forward = True
            find.Wrap = 0 # wdFindStop
            if find.Execute():
                # Information(1) = wdActiveEndAdjustedPageNumber (page number within section)
                # Information(3) = wdActiveEndPageNumber (absolute page number from beginning of document)
                sec_pg = rng.Information(1)
                abs_pg = rng.Information(3)
                print(f"  {p_name:25}: sec_pg={sec_pg}, abs_pg={abs_pg}")
            else:
                print(f"  {p_name:25}: NOT FOUND")

        # Check pages in Section 3 using Find
        print("\n--- SECTION 3 CHAPTERS & SECTIONS ---")
        targets = [
            ("Chapter 1", "Chapter 1"),
            ("1.1", "1.1 Background of the Project Topic"),
            ("1.2", "1.2 Motivation and Scope of the Report"),
            ("1.3", "1.3 Problem Statement"),
            ("1.4", "1.4 Salient Contributions"),
            ("1.5", "1.5 Organization of the Report"),
            ("Chapter 2", "Chapter 2"),
            ("2.1", "2.1 Introduction to the Literature Survey"),
            ("2.2", "2.2 Detailed Review of Existing Research"),
            ("Chapter 3", "Chapter 3"),
            ("3.1", "3.1 Block Diagram and Architectural Framework"),
            ("3.2", "3.2 Hardware Description"),
            ("3.3", "3.3 Software Description, Flowchart and Algorithms"),
            ("3.4", "3.4 Technical Implementation Challenges and Solutions"),
            ("Chapter 4", "Chapter 4"),
            ("4.1", "4.1 System Verification and Functional Test Matrix"),
            ("4.2", "4.2 Qualitative Demonstration and Live Session Outcomes"),
            ("4.3", "4.3 Latency, Resource Consumption, and Performance Analysis"),
            ("4.4", "4.4 Comparative Evaluation of Architecture Trade-offs"),
            ("4.5", "4.5 Project Development Timeline and Milestone Tracking"),
            ("Chapter 5", "Chapter 5"),
            ("5.1", "5.1 Advantages"),
            ("5.2", "5.2 Limitations"),
            ("5.3", "5.3 Applications"),
            ("Chapter 6", "Chapter 6"),
            ("6.1", "6.1 Conclusion"),
            ("6.2", "6.2 Future Scope"),
            ("References", "References"),
            ("Appendix I", "Appendix I: Sample Code Excerpts"),
            ("Appendix II", "Appendix II: Technical Specifications and Component Data Sheets"),
            ("Appendix III", "Appendix III: Comprehensive List of Components"),
            ("Appendix IV", "Appendix IV: List of Papers Presented and Published"),
        ]

        # For search in section 3, start searching after the first 100 characters of section 3 to avoid TOC if any
        for label, txt in targets:
            rng = s3.Range
            find = rng.Find
            find.ClearFormatting()
            find.Text = txt
            find.Forward = True
            find.Wrap = 0
            if find.Execute():
                sec_pg = rng.Information(1)
                abs_pg = rng.Information(3)
                print(f"  {label:20}: sec_pg={sec_pg}, abs_pg={abs_pg}")
            else:
                print(f"  {label:20}: NOT FOUND")

        print("\n--- FIGURES ---")
        figs = [
            ("Figure 3.1", "Figure 3.1: Complete End-to-End Five-Layer System Architecture"),
            ("Figure 3.2", "Figure 3.2: UML Use Case Diagram"),
            ("Figure 3.3", "Figure 3.3: UML Activity Diagram"),
            ("Figure 3.4", "Figure 3.4: UML Class Diagram"),
            ("Figure 3.5", "Figure 3.5: Entity-Relationship (ER) Diagram"),
            ("Figure 3.6", "Figure 3.6: UML Sequence Diagram"),
            ("Figure 3.7", "Figure 3.7: Component Ecosystem Architecture"),
            ("Figure 4.1", "Figure 4.1: Live Student Agent Operational Dashboard"),
            ("Figure 4.2", "Figure 4.2: Automated Active Google Meet Interaction"),
            ("Figure 4.3", "Figure 4.3: Project Development Timeline and Milestone Gantt Chart"),
        ]
        for label, txt in figs:
            rng = s3.Range
            find = rng.Find
            find.ClearFormatting()
            find.Text = txt
            find.Forward = True
            find.Wrap = 0
            if find.Execute():
                sec_pg = rng.Information(1)
                abs_pg = rng.Information(3)
                print(f"  {label:20}: sec_pg={sec_pg}, abs_pg={abs_pg}")
            else:
                print(f"  {label:20}: NOT FOUND")

        print("\n--- TABLES ---")
        tbls = [
            ("Table 2.1", "Table 2.1: Comprehensive Comparative Summary"),
            ("Table 3.1", "Table 3.1: Hardware and Computational Environment"),
            ("Table 3.2", "Table 3.2: Audio Ingestion and Buffer"),
            ("Table 4.1", "Table 4.1: Comprehensive Functional Verification"),
            ("Table 4.2", "Table 4.2: Computational Resource Footprint"),
            ("Table 4.3", "Table 4.3: Economic Compute Cost Comparison"),
            ("Table 4.4", "Table 4.4: Architectural Trade-off Analysis"),
            ("Table B.1", "Table B.1: VOSK Kaldi Acoustic Model Specifications"),
            ("Table B.2", "Table B.2: OpenAI GPT-4o-mini Operational Parameter Specifications"),
            ("Table C.1", "Table C.1: Comprehensive Hardware Component Manifest"),
            ("Table C.2", "Table C.2: Python Package Ecosystem and Installed Dependency Manifest"),
        ]
        for label, txt in tbls:
            rng = s3.Range
            find = rng.Find
            find.ClearFormatting()
            find.Text = txt
            find.Forward = True
            find.Wrap = 0
            if find.Execute():
                sec_pg = rng.Information(1)
                abs_pg = rng.Information(3)
                print(f"  {label:20}: sec_pg={sec_pg}, abs_pg={abs_pg}")
            else:
                print(f"  {label:20}: NOT FOUND")

    finally:
        wdoc.Close(False)
        word.Quit()

if __name__ == "__main__":
    main()
