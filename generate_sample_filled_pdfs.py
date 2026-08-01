"""
Generate 10 Sample Filled AcroForm PDF Files for Testing Data Extraction
==========================================================================
This script generates an interactive PDF template form using 'fillable_pdf_configurator.xlsx',
and populates 10 distinct filled PDF forms with realistic sample progress data.

Author: Interactive PDF Form Compiler Pipeline
"""

import os
import sys
import datetime
from pypdf import PdfReader, PdfWriter

sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from pdf_generator_from_excel import generate_pdf_from_excel


def create_10_sample_filled_pdfs():
    excel_path = os.path.abspath("Interactive_PDF_Code/fillable_pdf_configurator.xlsx")
    output_dir = os.path.abspath("Interactive_PDF_Code/Interactive_PDF")
    os.makedirs(output_dir, exist_ok=True)

    # 1. Generate base empty fillable PDF template form
    print("[INFO] Generating base fillable PDF template...")
    template_pdf = generate_pdf_from_excel(excel_path)
    abs_template = os.path.abspath(template_pdf)
    print(f"[SUCCESS] Base PDF template created at: '{abs_template}'")

    # Sample Rigs & Datasets for 10 distinct filled PDF forms
    rig_names = ["Rig Alpha", "Rig Beta", "Rig Gamma", "Rig Delta", "Rig Epsilon", "Rig Alpha", "Rig Beta", "Rig Gamma", "Rig Delta", "Rig Epsilon"]
    well_ids = ["WELL_A_01", "WELL_B_02", "WELL_C_03", "WELL_D_04", "WELL_E_05", "WELL_F_06", "WELL_G_07", "WELL_H_08", "WELL_I_09", "WELL_J_10"]
    
    sample_files = []

    print("[INFO] Generating 10 sample filled PDF forms...")
    for i in range(1, 11):
        rig = rig_names[i - 1]
        well = well_ids[i - 1]
        day_str = f"{i:02d}-07-2026"
        opening_stock = 100.0 + (i * 10.0)
        req = 50.0 + (i * 5.0)
        cons = 15.0 + (i * 2.0)
        closing = opening_stock - cons
        swell = round(1.2 + (i * 0.25), 2)
        wow = "Yes" if i % 3 == 0 else "No"
        food_box_date = f"{max(1, i-2):02d}-06-2026"
        narrative = f"Sample Daily Progress Report #{i}: Completed 24-hr drilling operation on Well {well} with Rig {rig}. ROP avg 12.5 m/hr. Mud weight {1.15 + (i*0.01):.2f} SG."
        short_dpr = f"Rig: {rig} | Well: {well} | Date: {day_str} | HSD Bal: {closing} KL | Swell: {swell}m | WOW: {wow}"

        field_data = {
            "Field_1_1": rig,
            "Field_1_2": well,
            "Field_1_3": day_str,
            "Field_2_1": f"{opening_stock:.1f}",
            "Field_2_2": f"{req:.1f}",
            "Field_2_3": f"{cons:.1f}",
            "Field_2_4": f"{closing:.1f}",
            "Field_3_1": f"{swell:.2f}",
            "Field_3_2": wow,
            "Field_3_3": food_box_date,
            "Field_4_1": narrative,
            "Field_5_1": short_dpr
        }

        # Populate AcroForm field values using pypdf
        reader = PdfReader(abs_template)
        writer = PdfWriter()
        writer.append(reader)
        writer.update_page_form_field_values(writer.pages[0], field_data)

        out_pdf_name = f"Sample_Filled_DPR_{i:02d}_{rig.replace('/', '_')}_{datetime.datetime.now().strftime('%Y%m%d')}.pdf"
        out_pdf_path = os.path.join(output_dir, out_pdf_name)

        with open(out_pdf_path, "wb") as f_out:
            writer.write(f_out)

        sample_files.append(out_pdf_path)
        print(f"  [{i}/10] Created Sample Filled PDF: '{out_pdf_name}'")

    print(f"\n[SUCCESS] Successfully generated 10 Sample Filled PDF forms in: '{output_dir}'")
    return sample_files


if __name__ == "__main__":
    create_10_sample_filled_pdfs()
