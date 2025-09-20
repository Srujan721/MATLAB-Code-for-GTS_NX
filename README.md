# MATLAB-Code-for-GTS_NX
# Tunnel Lining Analysis Automation (GTS NX Post-Processing)

This MATLAB-based toolkit automates the extraction, organization, and evaluation of structural analysis results for tunnel linings, specifically post-processed from GTS NX software output files.

## Features

- **Automated Excel Pre-processing:** Injects and executes VBA macros to format and clean raw result files.
- **Demand-Capacity Ratio Calculation:** Computes DCRs for structural elements, based on interpolated capacity curves.
- **Damage State Classification:** Classifies lining segments into five states from 'None' to 'Collapse.'
- **Batch Processing:** Handles multiple input files with robust error handling and validation.
- **Structured Excel Output:** Produces clear, organized output for both bending moment and shear force analyses.

## Repository Structure

- `src/` - Main MATLAB source code files.
- `macros/` - VBA macro(s) for Excel file formatting.
- `data/input/` - Place GTS NX result files here.
- `data/output/` - Output Excel files generated here.
- `data/0) Capacity_Curves.xlsx` - Required for DCR calculations.
- `docs/` - Documentation and figures.
- `examples/` - Example scripts for running the workflow.

## Getting Started

### Prerequisites
- MATLAB (test version: R20XXx)
- Microsoft Excel (with macros enabled)
- GTS NX tunnel analysis output files (as `.xlsx`)
- Excel macro security set to allow programmatic access

### Instructions

1. **Prepare Data**
   - Place all GTS NX result Excel files in `data/input/`.
   - Ensure `0) Capacity_Curves.xlsx` is in the `data/` directory.

2. **Configure Macros**
   - See `/macros/data_cleanup_macro.bas` for the required VBA macro.
   - Macro security: Ensure Excel allows macros to run programmatically.

3. **Run the Script**
   - Set working directory to repo root in MATLAB.
   - Execute `src/main.m`.

4. **Retrieve Results**
   - Output Excel files will appear in `data/output/` with `_Output` suffix.

## Output Description

- `DCR_BM` - Bending moment results (element number, max force, capacity, DCR, damage state)
- `DCR_SF` - Shear force results (similarly structured)
- `Final_results` - Damage state summaries for each file/section

## Notes

- Skips files not matching expected format.
- Continues processing if a file fails.
- Assumes GTS NX specific column structure.

## Example

See `examples/usage_example.m` for usage guidance.

## License

[MIT](LICENSE)

## Contributors

- Your Name

## Acknowledgments

- Based on analysis methodology for Hill Bot-Q(0.01)-D(9) model.
