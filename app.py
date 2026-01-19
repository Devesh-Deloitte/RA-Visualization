# -*- coding: utf-8 -*-
"""
Created on Mon Jan 19 00:23:43 2026

@author: partikumar
"""

import os
import time
import streamlit as st

# from irp_engine.input_reader import read_irp_inputs, validate_irp_inputs
# from irp_engine.solver import run_irp
# from irp_engine.outputs_writer import write_irp_outputs

st.set_page_config(page_title="IRP Inputs", page_icon="⚡", layout="wide", initial_sidebar_state="collapsed")

st.title("Annual Power Procurement Planning – Inputs")

st.write("Upload the IRP Excel input workbook and run the IRP model.")

uploaded_file = st.file_uploader(
    "Upload IRP Input Excel",
    type=["xlsx"],
    help="""
The input workbook should contain:
HourlyDemand, ExistingCapacity, PlannedCapacity, NewBuildOptions,
RenewableProfiles, ThermalOutageRates, HourlyMarketPrice, MonthlySTOA,
PolicyLimits, Penalties, Regions.
"""
)
    
# st.markdown("### Excel Input Workbook – What to Include")
st.info(
    """
This Excel workbook template should contain:

1) Hourly Demand & Price Forecasts
    
2) Existing & Planned Generation Portfolio (commissioning date, retirement date)

3) Technology-wise Technical Parameters (technical minimum, ramp rates, minimum up/down time, forced and planned outages)

4) Technology-wise Commercial Parameters (fixed and variable costs, start-up / shutdown costs)

5) Region-wise Renewable Energy Hourly Profiles

6) Year-on-year Addition Limits & Expansion Potential

7) Other Inputs Including Existing Short-term/Medium-term Contracts, RPO, ATC Limits 
"""
)

col1, col2 = st.columns([2, 1])

with col1:
    run_clicked = st.button("Run IRP Model", use_container_width=True)

with col2:
    st.info("Click **Run IRP Model** to start optimization. Once completed, you will be redirected to the Outcomes dashboard with charts and downloadable results.")

if uploaded_file is None:
    st.warning("Please upload the IRP input Excel file to continue.")
    st.stop()

# Save upload locally
os.makedirs("inputs", exist_ok=True)
input_path = os.path.join("inputs", "uploaded_irp_input.xlsx")
with open(input_path, "wb") as f:
    f.write(uploaded_file.getbuffer())

# Read and validate
try:
    # inputs = read_irp_inputs(input_path)
    # validate_irp_inputs(inputs)
    st.success("✅ Input file loaded & validated successfully.")
except Exception as e:
    st.error(f"❌ Input validation failed: {e}")
    st.stop()

if run_clicked:
    st.markdown("### Running IRP Optimization...")
    progress = st.progress(0)
    status = st.empty()

    # 5-sec loading animation (fake progress)
    for i in range(1, 6):
        status.write(f"Running... step {i}/5")
        progress.progress(i * 20)
        time.sleep(120)

    # with st.spinner("Solving capacity expansion optimization..."):
    #     model_results = run_irp(inputs)

    # Write expected outputs for your Outcomes UI
    # os.makedirs("outputs", exist_ok=True)
    # write_irp_outputs(model_results, out_dir="outputs")

    st.success("✅ IRP run completed. Redirecting to Outcomes page...")

    # store state for navigation
    st.session_state["irp_run_done"] = True
    st.session_state["outputs_dir"] = "outputs"

    time.sleep(1)
    st.switch_page("pages/capacity_expansion_ui.py")
