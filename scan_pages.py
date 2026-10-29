import pypdf

reader = pypdf.PdfReader('AI_Student_Agent_Black_Book_Final.pdf')
items = [
    ('List of Figures', 'LIST OF FIGURES'),
    ('List of Tables', 'LIST OF TABLES'),
    ('Abbreviations', 'ABBREVIATIONS'),
    ('Chapter 1', '1.1 Background of the Project Topic'),
    ('1.1 Background', '1.1 Background of the Project Topic'),
    ('1.2 Motivation', '1.2 Motivation and Scope of the Report'),
    ('1.3 Problem', '1.3 Problem Statement'),
    ('1.4 Contributions', '1.4 Salient Contributions'),
    ('1.5 Organization', '1.5 Organization of the Report'),
    ('Chapter 2', '2.1 Introduction to the Literature Survey'),
    ('2.1 Intro', '2.1 Introduction to the Literature Survey'),
    ('2.2 Review', '2.2 Detailed Review of Existing Research'),
    ('Chapter 3', '3.1 Block Diagram and Architectural Framework'),
    ('3.1 Architecture', '3.1 Block Diagram and Architectural Framework'),
    ('3.2 Hardware', '3.2 Hardware Description'),
    ('3.3 Software', '3.3 Software Description, Flowchart and Algorithms'),
    ('3.4 Challenges', '3.4 Technical Implementation Challenges and Solutions'),
    ('Chapter 4', '4.1 System Verification and Functional Test Matrix'),
    ('4.1 Verification', '4.1 System Verification and Functional Test Matrix'),
    ('4.2 Demo', '4.2 Qualitative Demonstration and Live Session Outcomes'),
    ('4.3 Performance', '4.3 Latency, Resource Consumption, and Performance Analysis'),
    ('4.4 Trade-offs', '4.4 Comparative Evaluation of Architecture Trade-offs'),
    ('4.5 Timeline', '4.5 Project Development Timeline and Milestone Tracking'),
    ('Chapter 5', '5.1 Advantages'),
    ('5.1 Advantages', '5.1 Advantages'),
    ('5.2 Limitations', '5.2 Limitations'),
    ('5.3 Applications', '5.3 Applications'),
    ('Chapter 6', '6.1 Conclusion'),
    ('6.1 Conclusion', '6.1 Conclusion'),
    ('6.2 Future Scope', '6.2 Future Scope'),
    ('References', 'References'),
    ('Appendix I', 'Appendix I: Sample Code Excerpts'),
    ('Appendix II', 'Appendix II: Technical Specifications and Component Data Sheets'),
    ('Appendix III', 'Appendix III: Comprehensive List of Components'),
    ('Appendix IV', 'Appendix IV: List of Papers Presented and Published'),
    ('Figure 3.1', 'Figure 3.1: Complete End-to-End Five-Layer System Architecture'),
    ('Figure 3.2', 'Figure 3.2: UML Use Case Diagram'),
    ('Figure 3.3', 'Figure 3.3: UML Activity Diagram'),
    ('Figure 3.4', 'Figure 3.4: UML Class Diagram'),
    ('Figure 3.5', 'Figure 3.5: Entity-Relationship (ER) Diagram'),
    ('Figure 3.6', 'Figure 3.6: UML Sequence Diagram'),
    ('Figure 3.7', 'Figure 3.7: Component Ecosystem Architecture'),
    ('Figure 4.1', 'Figure 4.1: Live Student Agent Operational Dashboard'),
    ('Figure 4.2', 'Figure 4.2: Automated Active Google Meet Interaction'),
    ('Figure 4.3', 'Figure 4.3: Project Development Timeline and Milestone Gantt Chart'),
    ('Table 2.1', 'Table 2.1: Comprehensive Comparative Summary of Reviewed Literature'),
    ('Table 3.1', 'Table 3.1: Hardware and Computational Environment Specifications'),
    ('Table 3.2', 'Table 3.2: Audio Ingestion and Buffer Dimension Parameters'),
    ('Table 4.1', 'Table 4.1: Comprehensive Functional Verification and Test Matrix'),
    ('Table 4.2', 'Table 4.2: Computational Resource Footprint and Latency Metrics'),
    ('Table 4.3', 'Table 4.3: Economic Compute Cost Comparison: Continuous Cloud vs Hybrid Local'),
    ('Table 4.4', 'Table 4.4: Architectural Trade-off Analysis Across Alternative Implementations'),
    ('Table B.1', 'Table B.1: VOSK Kaldi Acoustic Model Specifications'),
    ('Table B.2', 'Table B.2: OpenAI GPT-4o-mini Operational Parameter Specifications'),
    ('Table C.1', 'Table C.1: Comprehensive Hardware Component Manifest'),
    ('Table C.2', 'Table C.2: Python Package Ecosystem and Installed Dependency Manifest'),
]

# Cache page texts first
page_texts = [page.extract_text() for page in reader.pages]

with open('pages_map.txt', 'w', encoding='utf-8') as f:
    f.write(f'TOTAL PAGES: {len(page_texts)}\n')
    for label, term in items:
        found = False
        term_clean = ' '.join(term.lower().split())
        start_idx = 15 if not label.startswith('List') and not label.startswith('Abbr') else 0
        for p_idx in range(start_idx, len(page_texts)):
            # If prelim, only search in first 15 pages
            if start_idx == 0 and p_idx >= 15:
                break
            txt_clean = ' '.join(page_texts[p_idx].lower().split())
            if term_clean in txt_clean:
                sec3_pg = p_idx - 15 + 1
                abs_pg = p_idx + 1
                f.write(f'{label:22} | Sec3 Page: {sec3_pg:2d} | Abs Page: {abs_pg:2d}\n')
                found = True
                break
        if not found:
            f.write(f'{label:22} | NOT FOUND\n')

print("DONE writing pages_map.txt", flush=True)
