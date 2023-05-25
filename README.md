# DSF Tm Automation
 A Python automation script for deriving and storing melting point (Tm) data from Differential Scanning Fluorimetry (DSF) first-derivative plots for protein-small molecule binding.

Differential scanning fluorimetry (DSF) is a biophysical "thermal shift" technique used for examining how a protein unfolds when it is steadily heated over a controlled temperature range. One of its applications is the detection of ligands (which can be small molecules or other proteins) that bind to the target protein. As the ligand binds to the protein, it changes the _melting point_ (T~m~) of the protein, defined as the temperature at which 50% of the protein is in the folded state and 50% is in the unfolded state. This shift in T~m (∆T~m~) varies based on many factors, including the structure and concentration of the ligand ([Gao et al. 2020](https://link.springer.com/article/10.1007/s12551-020-00619-2)).

The Python script in this repository was developed as part of a research project which involved DSF assays, where various different small molecule ligands at different concentrations were heated with a target protein, and the resulting T~m~ shifts were used to deduce the binding mechanism of the ligands to the protein. Note that instead of raw response-vs-temperature plots, this script works with first-derivative plots (of response with respect to temperature) for the sake of convenience.

![DSF plot generated from the first-derivative data, including all the melting points captured by the script within the bounds set by the user.](https://raw.githubusercontent.com/RaiyanR86/DSF_Tm_Automation/main/TmPlot_Compound_001.png)

Steps for using the code:
1. Download the ![DSF_Tm_automation](https://github.com/RaiyanR86/DSF_Tm_Automation/blob/main/DSF_Tm_automation.py) script.
2. (Optional) Downloading the ![Input_Dataset.xlsx](https://github.com/RaiyanR86/DSF_Tm_Automation/blob/main/Input_Dataset.xlsx) and ![Output_File.xlsx](https://github.com/RaiyanR86/DSF_Tm_Automation/blob/main/Output_File.xlsx) files will provide help in formatting the user data. (Make sure source workbook cells are values only and not formulas.)
3. In lines 20 and 25 of the script, type in input and output file names in the requested format. Make sure input file is in same folder as Python script.
4. Run program and input all variables as requested ("recommended" variables are for Input_Dataset only). (NOTE: If destination file is already open, please close it before running the program, otherwise program will not be able to access destination file.)
