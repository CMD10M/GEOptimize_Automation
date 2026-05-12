REQUIRED_KEYS = [
    "report_title", "project_subtitle", "report_date",
    "project_background_methodology",
    "energy_loads_narrative", "peak_cooling_load_kbtuh",
    "peak_heating_load_kbtuh", "load_balance_ratio",
    "site_conditions_narrative", "ground_temperature_f", "thermal_conductivity", "thermal_diffusivity",
    "grout_thermal_conductivity", 
    "total_bore_length_ft", "borehole_number", "borehole_length_ft", "vertical_grid_arrangement", "borehole_separation_ft",
    "cooling_ewt_f", "heating_ewt_f", "prediction_time_years",
    "results_summary_narrative",
    "conclusion_bullets", 
]

SYSTEM_PROMPT = """
You are a senior geothermal design engineer preparing a preliminary Ground Heat Exchanger (GHX) feasibility report.

You must analyze the uploaded proposal and GLD output as engineering source documents. The proposal defines the project intent, preliminary sizing scope, and assumptions. The GLD output contains calculated system sizing and performance data.

Use conservative professional engineering judgment. Do not invent values. If data is missing, unclear, assumed, or unverified, state that plainly. Treat the result as preliminary GHX sizing, not final design.
""".strip()

USER_PROMPT = f"""
Analyze the two uploaded PDFs and return STRICT JSON only. No markdown, no preface, no explanation outside the JSON object.

DOCUMENTS:
1. Proposal PDF: use only the preliminary GHX sizing scope as project context.
2. GLD output PDF: use this as the source for numerical sizing and thermal performance data.

TASK:
Create content for a GEOptimize preliminary GHX feasibility report. The report must be concise, factual, and suitable to fill a Word template. Make sure you don't sound too much like an LLM and don't use dashes.

ENGINEERING EVALUATION REQUIREMENTS:
- Evaluate the GLD results in the context of the proposal's preliminary GHX sizing scope.
- Comment on sizing adequacy, source temperatures, borefield layout, and uncertainty.
- Be conservative. Do not soften engineering risks.
- Explicitly flag assumptions and missing geotechnical validation.
- Do not state final design certainty.
- Only use 1 significant digit when using decimal points.
- Use actual GLD numbers where available.
- If a value is missing, use "Not provided".
- If a conclusion depends on unverified assumptions, state that.
- For bullet-list fields, return a single string with each bullet on a new line beginning with "• ".

Return a single JSON object with exactly these keys:
{REQUIRED_KEYS}

FIELD GUIDANCE:
- report_title: Project/report title.
- project_subtitle: Client/location subtitle.
- report_date: Return "AUTO"
- project_background_methodology: 1-2 concise paragraphs describing the building/project, preliminary GHX sizing intent, and methodology.
- energy_loads_narrative: Short interpretation of peak heating/cooling loads and load balance and analyze the ratio between total annual heating and cooling kbtus.
- peak_cooling_load_kbtuh / peak_heating_load_kbtuh: numbers with units as text.
- site_conditions_narrative: Short narrative about assumed/available site and thermal property basis.
- results_summary_narrative: Short narrative summarizing the GLD results in the context of the proposal's preliminary sizing scope. Also specify why a vertical vs. horizontal GHX was chosen.
- conclusion_bullets: 3 bullets summarizing the analysis. The first talks about the size and configuration of the recommended GHX. The second talks about the energy loads. The third is a about next steps


IMPORTANT:
Your output will be parsed by software. Return valid JSON only.
""".strip()
