# ATM
Version 1.0

Automated Trace Maker = ATM

To quickly prepare traces of reversed-phase high-performance liquid chromatography (RP-HPLC) experiments, three different ATM Python scripts were developed. The standard ATM program takes .CSV RP-HPLC files (containing only retention times and absorbance values) and creates both raw and normalized Excel plots of the data. The MS_ATM (Mass Spectra Automated Trace Maker) program was developed to analyze mass spectra data generated from liquid chromatography-mass spectrometry (LC-MS) experiments. MS_ATM takes .CSV mass spectra files (containing only m/z values and relative intensities) and creates Excel plots of the data. Finally, the SUPER_ATM (SUPERimposed Automated Trace Maker) program creates superimposed plots of all RP-HPLC data within a folder. SUPER_ATM takes .CSV files (containing only retention times and absorbances values) to create superimposed raw and normalized Excel plots of the data.

All three ATM programs are available as executable files. To use the executables, first download the appropriate version based on your computer's operating system (go to the "Releases" tab to find the latest executable files). After downloading the executable, place it into a folder containing the RP-HPLC or LC-MS .CSV files that you would like to plot. Click on the executable file to start the ATM program, follow any on-screen instructions to enter user inputs, and your Excel plots will appear in the same folder that contains the exectuable and .CSV files.

If you're using MS_ATM, please note that the program assumes that the minimum m/z value is 400 and that the maximum m/z value is 2,000. If your experiment had different minimum and maximum values, please change the plot's x-axis within Excel to see your data.

All three ATM scripts are compatible with Python 2.7. If you're using the Python scripts manually, please note that the openpyxl library must be installed for these scripts to work.
