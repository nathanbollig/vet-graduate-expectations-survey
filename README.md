# vet-graduate-expectations-survey

## Analysis steps

1. Move data into working directory (not included in this repo)
2. Run `analysis.ipynb`
    - Within the notebook, alter `analysis_mode` and `nontechnical` flag to perform each of the analysis
    - This generates primary results files with the file naming suffixes as defined in the notebook (only analysis mode 0 and 2 are used for reporting in the manuscript)
3. Run `make_tables.py --top n` with desired value of n to produce some of the data tables that will be included in the manuscript.
4. Use `WVMA_table.ipynb` to generate the table of tasks most expected by WMVA members to be performed independently. 
