import os
import sys
import time
import win32com.client

# Ensure unbuffered standard output
sys.stdout.reconfigure(line_buffering=True)

def post_process_report():
    print("=" * 65, flush=True)
    print(" Post-Processing Black Book Document via Microsoft Word COM...", flush=True)
    print("=" * 65, flush=True)

    base_dir = r"C:\Users\ganesh jadhav\Downloads\CAPSTONE\AI-Student-Agent"
    doc_path = os.path.join(base_dir, "AI_Student_Agent_Black_Book_Final.docx")
    pdf_path = os.path.join(base_dir, "AI_Student_Agent_Black_Book_Final.pdf")

    if not os.path.exists(doc_path):
        raise FileNotFoundError(f"Document not found: {doc_path}")

    print(" -> Connecting to Microsoft Word COM...", flush=True)
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    wdoc = None
    try:
        print(f" -> Opening document: {doc_path} ...", flush=True)
        t0 = time.time()
        wdoc = word.Documents.Open(doc_path)
        print(f" -> Document loaded in {time.time()-t0:.2f}s.", flush=True)
        print(f" -> Section Count: {wdoc.Sections.Count}", flush=True)

        # 1. Configure Section 1 (Cover Page & Certificate)
        s1 = wdoc.Sections(1)
        s1.Headers(1).Range.Text = ""
        s1.Footers(1).Range.Text = ""
        print(" -> Section 1: Headers and Footers cleared (no page numbers).", flush=True)

        # 2. Configure Section 2 (Preliminary Pages)
        s2 = wdoc.Sections(2)
        s2.Headers(1).LinkToPrevious = False
        s2.Headers(1).Range.Text = ""
        s2.Footers(1).LinkToPrevious = False
        s2.Footers(1).PageNumbers.RestartNumberingAtSection = True
        s2.Footers(1).PageNumbers.StartingNumber = 1
        s2.Footers(1).PageNumbers.NumberStyle = 2  # wdPageNumberStyleLowercaseRoman = 2
        s2.Footers(1).PageNumbers.Add(2)          # wdAlignPageNumberRight = 2
        print(" -> Section 2: Configured Lowercase Roman page numbering starting at 'i'.", flush=True)

        # 3. Configure Section 3 (Main Chapters onward)
        s3 = wdoc.Sections(3)
        s3.Headers(1).LinkToPrevious = False
        s3.Headers(1).Range.Text = "AI Student Agent: An Autonomous AI Learning Assistant for Real-Time Online Classes\t\tAcademic Year 2026-2027"
        s3.Headers(1).Range.Font.Name = "Times New Roman"
        s3.Headers(1).Range.Font.Size = 10
        s3.Footers(1).LinkToPrevious = False
        s3.Footers(1).PageNumbers.RestartNumberingAtSection = True
        s3.Footers(1).PageNumbers.StartingNumber = 1
        s3.Footers(1).PageNumbers.NumberStyle = 0  # wdPageNumberStyleArabic = 0
        s3.Footers(1).PageNumbers.Add(2)          # wdAlignPageNumberRight = 2
        print(" -> Section 3: Configured Arabic decimal page numbering starting at '1', with official institutional running header.", flush=True)

        # Force repagination
        print(" -> Repaginating document across all sections...", flush=True)
        wdoc.Repaginate()
        print(" -> Repaginated document successfully.", flush=True)

        # Save finalized DOCX
        print(" -> Saving configured DOCX...", flush=True)
        wdoc.Save()
        print(f" -> Saved updated DOCX: '{doc_path}'", flush=True)

        # Export to PDF natively via Word
        print(" -> Exporting document to PDF (Format 17)...", flush=True)
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except Exception:
                pass

        t_pdf = time.time()
        wdoc.SaveAs2(pdf_path, FileFormat=17)
        print(f" [SUCCESS] Successfully exported PDF in {time.time()-t_pdf:.2f}s: '{pdf_path}'", flush=True)
        print(f" PDF File Size: {os.path.getsize(pdf_path):,} bytes", flush=True)
        print(f" Total Page Count: {wdoc.ComputeStatistics(2)}", flush=True)
        print("=" * 65, flush=True)

    finally:
        if wdoc is not None:
            try:
                wdoc.Close(False)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass

if __name__ == "__main__":
    post_process_report()
