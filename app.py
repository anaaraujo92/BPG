import streamlit as st
import pandas as pd
from fpdf import FPDF

# Function to calculate dilution and weighing
def calculate_preparation(target_concentration, final_volume, weighing_range):
    ideal_weight = target_concentration * final_volume  # Calculate required mass
    min_weight, max_weight = weighing_range  # Allowed weighing range
    
    return ideal_weight, min_weight, max_weight

# Function to determine preparation based on validation parameters
def get_preparation_details(validation_parameters):
    details = {
        "Selectivity": "Prepare blank solution, standard solution, and sample solutions with impurities spiked at their MLD levels.\n"
                      "Ensure no interference is observed between the blank, excipients, and peaks of interest.",
        "Accuracy": "Prepare 'as is' sample solutions and three independent samples spiked with impurities at LOQ, target concentration, "
                    "and 120% of the maximum defined limit. Maintain excipients at target concentration.",
        "Repeatability": "Prepare six independent sample solutions at target concentration. If no quantifiable impurities are observed, "
                        "spike samples with impurities at their maximum limit.",
        "Stability of Solutions": "Inject the sample solution immediately after preparation (T0) and at intervals (12h, 24h, 48h) to assess stability. "
                                 "Spiked samples should be used if no quantifiable impurities are observed.",
        "Linearity": "Prepare standard solutions at five different concentrations, covering 50% to 150% of the method's target concentration.",
        "Robustness": "Assess variations in analytical conditions, such as column temperature, flow rate, and mobile phase composition."
    }
    return "\n".join([details.get(param, "No details available.") for param in validation_parameters])

# Function to generate PDF report
def generate_pdf(description):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, "Bench Plan Report", ln=True, align="C")
    pdf.ln(10)
    pdf.multi_cell(0, 10, description)
    output_pdf = "Generated_Bench_Plan.pdf"
    pdf.output(output_pdf)
    return output_pdf

# Streamlit interface
st.title("Bench Plan Generator for Method Validation")

# User inputs
validation_parameters = st.multiselect("Select Validation Parameters:", ["Selectivity", "Accuracy", "Repeatability", "Stability of Solutions", "Linearity", "Robustness"])
solution_type = st.selectbox("Select Solution Type:", ["Sample As Is", "Spiked Sample", "Standard", "Placebo"])
target_concentration = st.number_input("Enter Target Concentration (mg/mL):", min_value=0.0, format="%.4f")
final_volume = st.number_input("Enter Final Volume (mL):", min_value=0.0, format="%.2f")
solvent = st.text_input("Enter Solvent:")

# Define allowed weighing range (example: 90% to 110% of the ideal value)
weighing_range = (target_concentration * 0.90 * final_volume,
                  target_concentration * 1.10 * final_volume)

# Generate preparation details
if st.button("Generate Bench Plan"):
    ideal_weight, min_weight, max_weight = calculate_preparation(target_concentration, final_volume, weighing_range)
    validation_details = get_preparation_details(validation_parameters)
    
    description = (f"For {solution_type}, weigh {ideal_weight:.2f} mg "
                   f"(allowed range: {min_weight:.2f} mg to {max_weight:.2f} mg) "
                   f"and dissolve in {final_volume:.2f} mL of {solvent}.\n"
                   f"\n{validation_details}")
    
    # Create DataFrame
    bench_plan = pd.DataFrame([
        {"Step": f"Prepare {solution_type}", "Description": description, "Observations": ""}
    ])
    
    # Save to Excel
    output_excel = "Generated_Bench_Plan.xlsx"
    bench_plan.to_excel(output_excel, index=False)
    
    # Generate PDF
    output_pdf = generate_pdf(description)
    
    # Download buttons
    with open(output_excel, "rb") as f:
        st.download_button(
            label="📥 Download Bench Plan (Excel)",
            data=f,
            file_name="Generated_Bench_Plan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with open(output_pdf, "rb") as f:
        st.download_button(
            label="📥 Download Bench Plan (PDF)",
            data=f,
            file_name="Generated_Bench_Plan.pdf",
            mime="application/pdf"
        )
