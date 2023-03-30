# flow

This project models prospective material flows based on the [Open Dynamic Material Systems Model (ODYM)](https://github.com/indecol/odym). It currently includes:
- The ODYM model framework in "ODYM-Master";
- A prospective, stock-driven model of steel and concrete in buildings in the EU-27 in "buildings_pro_stock_EU".

## Requirements
The requirements for the model is stated in the file "requirements.txt". To import ODYM directly as a package, the directory must be installed (pip install .)

## Structure
The input data for the model can be found under "docs" in the model folder. The classification of flows and stocks can be found in the classification file and the configuration of the model can be found in the config file in the same folder. The results are available under "results" in the model folder. The underlying calculation is located under "src" in the model folder.

## Publications and further information
More information on the models will be available here:
- Paper on the building model (coming soon).

## Contact
meta.thurid.lotz@isi.fraunhofer.de
