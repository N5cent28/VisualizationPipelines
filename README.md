---

# Visualization Pipelines

Exhibiting some of the analysis -> visualization pipelines I have created.

## Overview

This repository contains various Jupyter Notebook files and Excel VBA scripts used to process and visualize data for different projects. As a bioinformatic scientist at Proteovista LLC, I have developed these pipelines to assist in the analysis and visualization of DNA microarray data.

## Contents

### [Protein_Comparison_nod2-11pt2-773.ipynb](https://github.com/N5cent28/VisualizationPipelines/blob/main/Protein_Comparison_nod2-11pt2-773.ipynb)

This Jupyter Notebook compares DNA microarray intensity data from three different proteins added to the array, examining their differential binding profiles. The notebook includes functions to:
- Convert specified columns to Z-scores.
- Filter outliers based on Z-scores.
- Plot histograms and KDE plots for the protein data.

### [ReplicateVenDiagram_vis.ipynb](https://github.com/N5cent28/VisualizationPipelines/blob/main/ReplicateVenDiagram_vis.ipynb)

This notebook helps quantify the replicability of array runs or determine which aptamers may have preferential binding to a protein of interest. It includes:
- Loading and processing data.
- Calculating top feature numbers for Venn diagram plotting.
- Creating Venn diagrams to visualize overlapping features.

### [UWsub_groupgraphing](https://github.com/N5cent28/VisualizationPipelines/blob/main/UWsub_groupgraphing)

An Excel VBA pipeline used to efficiently process DNA microarray data from Agilent feature extraction output to graphs of median feature intensity by feature group number. The script includes:
- Deleting unnecessary rows and columns.
- Adding headers and freezing panes.
- Sorting data and filtering out unwanted values.
- Extracting group numbers and calculating summary statistics for graphing.
- Creating and formatting multiple graphs.

### [VenDiagram&BoxWhisker_vis.ipynb](https://github.com/N5cent28/VisualizationPipelines/blob/main/VenDiagram&BoxWhisker_vis.ipynb)

This notebook aims to confirm replicability between array runs and determine if a change in buffer protocol altered protein binding results. It includes:
- Loading and processing data.
- Creating Venn diagrams to visualize overlapping features.
- Calculating and visualizing Spearman rank correlation between different experimental conditions.
- Plotting box and whisker plots to compare intensity distributions across replicates.

## Usage

1. Clone the repository:
   ```bash
   git clone https://github.com/N5cent28/VisualizationPipelines.git
   ```

2. Navigate to the project directory:
   ```bash
   cd VisualizationPipelines
   ```

3. Open the Jupyter Notebooks using JupyterLab or Jupyter Notebook:
   ```bash
   jupyter lab
   # or
   jupyter notebook
   ```

4. Execute the cells in the notebook to run the analysis and visualize the data.

For the Excel VBA script, follow these steps:
1. Open the `UWsub_groupgraphing` script in Excel.
2. Run the macros to process the data and create graphs.

## License

This repository is licensed under the MIT License. See the [LICENSE](LICENSE) file for more information.

---
