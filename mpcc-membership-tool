import streamlit as st
import xlwings as xw
from pathlib import Path

st.set_page_config(page_title="MPCC Membership Planning Tool", layout="wide")

EXCEL_FILE = Path("/Users/samquarles/Desktop/MPCC Membership MODEL.xlsx")

st.title("MPCC Membership Planning Tool")
st.write("Update member mix and assumptions, then run the scenario to see utilization and staffing results.")

if not EXCEL_FILE.exists():
    st.error(f"Excel file not found: {EXCEL_FILE}")
    st.stop()


def write_inputs(inputs_sheet, values):
    for cell, value in values.items():
        inputs_sheet[cell].value = value


def run_model(values):
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        wb = app.books.open(str(EXCEL_FILE))
        inputs = wb.sheets["Inputs"]
        dashboard = wb.sheets["Dashboard"]

        write_inputs(inputs, values)

        wb.app.calculate()

        outputs = {
            "ending_members": dashboard["C4"].value,
            "peak_rounds": dashboard["C5"].value,
            "peak_golf_capacity": dashboard["C6"].value,
            "golf_congestion": dashboard["C7"].value,
            "peak_dining_covers": dashboard["C8"].value,
            "peak_dining_capacity": dashboard["C9"].value,
            "dining_pressure": dashboard["C10"].value,
            "peak_daily_rounds_per_staff": dashboard["C11"].value,
            "peak_daily_covers_per_staff": dashboard["C12"].value,
            "golf_staffing_strain": dashboard["C13"].value,
            "dining_staffing_strain": dashboard["C14"].value,
            "scenario_status": dashboard["C16"].value,
        }

        wb.save()
        wb.close()
        return outputs

    finally:
        app.quit()


with st.sidebar:
    st.header("Workbook")
    st.caption(f"Using: {EXCEL_FILE.name}")

tab1, tab2 = st.tabs(["Scenario Inputs", "Assumptions"])

with tab1:
    st.subheader("Membership Scenario")

    c1, c2 = st.columns(2)

    with c1:
        lifestyle_count = st.number_input("Lifestyle count", min_value=0.0, value=700.0, step=1.0)
        golf_count = st.number_input("Golf count", min_value=0.0, value=111.0, step=1.0)
        occasional_count = st.number_input("Occasional count", min_value=0.0, value=167.0, step=1.0)

    with c2:
        social_count = st.number_input("Social count", min_value=0.0, value=120.0, step=1.0)
        inactive_count = st.number_input("Inactive count", min_value=0.0, value=50.0, step=1.0)
        annual_admissions = st.number_input("Annual admissions", min_value=0.0, value=36.0, step=1.0)

with tab2:
    st.subheader("Editable Assumptions")

    st.markdown("**Attrition assumptions**")
    a1, a2, a3, a4, a5 = st.columns(5)
    with a1:
        lifestyle_attrition = st.number_input("Lifestyle attrition", min_value=0.0, max_value=1.0, value=0.007, step=0.001, format="%.3f")
    with a2:
        golf_attrition = st.number_input("Golf attrition", min_value=0.0, max_value=1.0, value=0.007, step=0.001, format="%.3f")
    with a3:
        occasional_attrition = st.number_input("Occasional attrition", min_value=0.0, max_value=1.0, value=0.007, step=0.001, format="%.3f")
    with a4:
        social_attrition = st.number_input("Social attrition", min_value=0.0, max_value=1.0, value=0.007, step=0.001, format="%.3f")
    with a5:
        inactive_attrition = st.number_input("Inactive attrition", min_value=0.0, max_value=1.0, value=0.007, step=0.001, format="%.3f")

    st.markdown("**Golf usage assumptions**")
    g1, g2, g3 = st.columns(3)
    with g1:
        rounds_lifestyle = st.number_input("Rounds/member/year - Lifestyle", min_value=0.0, value=43.0, step=1.0)
        peak_rounds_lifestyle = st.number_input("Peak round share - Lifestyle", min_value=0.0, max_value=1.0, value=0.40, step=0.01, format="%.2f")
        weekend_share_lifestyle = st.number_input("Weekend round share - Lifestyle", min_value=0.0, max_value=1.0, value=0.40, step=0.01, format="%.2f")
    with g2:
        rounds_golf = st.number_input("Rounds/member/year - Golf", min_value=0.0, value=28.0, step=1.0)
        peak_rounds_golf = st.number_input("Peak round share - Golf", min_value=0.0, max_value=1.0, value=0.75, step=0.01, format="%.2f")
        weekend_share_golf = st.number_input("Weekend round share - Golf", min_value=0.0, max_value=1.0, value=0.55, step=0.01, format="%.2f")
    with g3:
        rounds_occasional = st.number_input("Rounds/member/year - Occasional", min_value=0.0, value=3.4, step=0.1)
        peak_rounds_occasional = st.number_input("Peak round share - Occasional", min_value=0.0, max_value=1.0, value=0.90, step=0.01, format="%.2f")
        weekend_share_occasional = st.number_input("Weekend round share - Occasional", min_value=0.0, max_value=1.0, value=0.65, step=0.01, format="%.2f")

    st.markdown("**Dining usage assumptions**")
    d1, d2 = st.columns(2)
    with d1:
        dining_lifestyle = st.number_input("Dining visits/member/year - Lifestyle", min_value=0.0, value=36.0, step=1.0)
        dining_golf = st.number_input("Dining visits/member/year - Golf", min_value=0.0, value=12.0, step=1.0)
        dining_occasional = st.number_input("Dining visits/member/year - Occasional", min_value=0.0, value=8.0, step=1.0)
    with d2:
        dining_social = st.number_input("Dining visits/member/year - Social", min_value=0.0, value=24.0, step=1.0)
        peak_dining_lifestyle = st.number_input("Peak dining share - Lifestyle", min_value=0.0, max_value=1.0, value=0.25, step=0.01, format="%.2f")
        peak_dining_golf = st.number_input("Peak dining share - Golf", min_value=0.0, max_value=1.0, value=0.40, step=0.01, format="%.2f")

    d3, d4 = st.columns(2)
    with d3:
        peak_dining_occasional = st.number_input("Peak dining share - Occasional", min_value=0.0, max_value=1.0, value=0.80, step=0.01, format="%.2f")
        peak_dining_social = st.number_input("Peak dining share - Social", min_value=0.0, max_value=1.0, value=0.50, step=0.01, format="%.2f")
    with d4:
        peak_night_share = st.number_input("Peak dining visits share on Thu-Sat", min_value=0.0, max_value=1.0, value=0.75, step=0.01, format="%.2f")

    st.markdown("**Capacity and staffing assumptions**")
    p1, p2, p3 = st.columns(3)
    with p1:
        tee_interval = st.number_input("Tee interval (minutes)", min_value=1.0, value=10.0, step=1.0)
        playable_hours = st.number_input("Playable hours per day", min_value=1.0, value=11.0, step=1.0)
        golfers_per_tee_time = st.number_input("Golfers per tee time", min_value=1.0, value=4.0, step=1.0)
    with p2:
        peak_golf_days = st.number_input("Peak golf days per year", min_value=1.0, value=90.0, step=1.0)
        realistic_capacity_factor = st.number_input("Realistic golf capacity factor", min_value=0.0, max_value=1.0, value=0.45, step=0.01, format="%.2f")
        dining_seats = st.number_input("Total dining seats", min_value=0.0, value=230.0, step=1.0)
    with p3:
        table_turns = st.number_input("Peak table turns per night", min_value=0.0, value=2.0, step=0.1)
        peak_dining_days = st.number_input("Peak dining days per year", min_value=1.0, value=90.0, step=1.0)
        golf_headcount = st.number_input("Golf operations headcount", min_value=0.0, value=29.0, step=1.0)

    s1, s2 = st.columns(2)
    with s1:
        fb_headcount = st.number_input("F&B headcount", min_value=0.0, value=50.0, step=1.0)
    with s2:
        rounds_per_staff_capacity = st.number_input("Rounds per golf staff capacity", min_value=0.0, value=20.0, step=1.0)
        covers_per_staff_capacity = st.number_input("Covers per F&B staff capacity", min_value=0.0, value=20.0, step=1.0)

input_map = {
    "C6": lifestyle_count,
    "C7": golf_count,
    "C8": occasional_count,
    "C9": social_count,
    "C10": inactive_count,
    "C12": annual_admissions,

    "C16": lifestyle_attrition,
    "C17": golf_attrition,
    "C18": occasional_attrition,
    "C19": social_attrition,
    "C20": inactive_attrition,

    "C24": rounds_lifestyle,
    "C25": rounds_golf,
    "C26": rounds_occasional,
    "C29": peak_rounds_lifestyle,
    "C30": peak_rounds_golf,
    "C31": peak_rounds_occasional,
    "C34": weekend_share_lifestyle,
    "C35": weekend_share_golf,
    "C36": weekend_share_occasional,

    "C39": dining_lifestyle,
    "C40": dining_golf,
    "C41": dining_occasional,
    "C42": dining_social,
    "C44": peak_dining_lifestyle,
    "C45": peak_dining_golf,
    "C46": peak_dining_occasional,
    "C47": peak_dining_social,

    "C51": tee_interval,
    "C52": playable_hours,
    "C53": golfers_per_tee_time,
    "C55": peak_golf_days,
    "C56": realistic_capacity_factor,
    "C62": dining_seats,
    "C63": table_turns,
    "C64": peak_dining_days,

    "C69": golf_headcount,
    "C70": fb_headcount,
    "C73": rounds_per_staff_capacity,
    "C74": covers_per_staff_capacity,
    "C77": peak_night_share,
}

if st.button("Run Scenario", type="primary"):
    try:
        results = run_model(input_map)

        st.subheader("Scenario Results")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Ending Members", f"{results['ending_members']:,.0f}")
        m2.metric("Peak Rounds", f"{results['peak_rounds']:,.0f}")
        m3.metric("Golf Congestion", f"{results['golf_congestion']:.1%}")
        m4.metric("Dining Pressure", f"{results['dining_pressure']:.1%}")

        m5, m6, m7, m8 = st.columns(4)
        m5.metric("Peak Golf Capacity", f"{results['peak_golf_capacity']:,.0f}")
        m6.metric("Peak Dining Covers", f"{results['peak_dining_covers']:,.0f}")
        m7.metric("Golf Staffing Strain", f"{results['golf_staffing_strain']:.1%}")
        m8.metric("Dining Staffing Strain", f"{results['dining_staffing_strain']:.1%}")

        st.markdown(f"### Scenario Status: **{results['scenario_status']}**")

        st.write("Additional staffing outputs")
        a1, a2 = st.columns(2)
        a1.metric("Peak Daily Rounds per Golf Staff", f"{results['peak_daily_rounds_per_staff']:,.2f}")
        a2.metric("Peak Daily Covers per F&B Staff", f"{results['peak_daily_covers_per_staff']:,.2f}")

    except Exception as e:
        st.error(f"Error running model: {e}")

