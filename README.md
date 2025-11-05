# Chemical Shift Perturbation (CSP) Histogram and Multi XY-scatter plotting 📊📈
Written by Uno Bolin Hartl (With chatGPT 5.0), BSc Candidate Molecular Biology (Stockholm University, MBW) and Research
assistant in Katja Petzold Lab (RNA Dynamics by NMR - Uppsala University). 

Histogram script (**Coords_v2**) offers plotting the ΔCSP per Y-scatter point.
Multi XY-scatter grid plot (**Multi_Choice**) offers plotting of 1-n numbers of XY-scatter plots with lines bridging points. 
The latter preferably followed by the script enlisted in **CSP_Fits_LinLogExpHill** repository.
All of these scripts require xlsx files, sometimes with some pre-processing. In **Multi_Choice** X and Y values have to be
on columns K and L respectively (To be patched/updated!).

This Script Is useful for:
- plotting virtually any change in value (Δ, designed for CSP) from an excel file, in to a histogram plot.
- plotting multiple XY-scatter plot in a grid.

---
##  Requirements 🚧
- **Python 3.8+**
- Install dependencies: **numpy, pandas, matplotlib and openpyxl**. Optionally numexpr for large datasets
- Alternatively use commands below for discription.
  ```bash
  pip install -r requirements.txt
---

## ✨ Key Features
| Feature | Description |
|--------|-------------|
| Works directly on **.xlsx** directories | No conversion required |
| Grid and Sorting (Coords_v2) | `Sorting by column first number` |
| Adjustable fig size | `e.g. 12, 8; width, height in inches` |
| Configurable Coloring  | `colormap or discrete color` |
| Adjust how many sets | How many grid-plots per png |
| input/output directory | Prevents overwriting unless allowed |
