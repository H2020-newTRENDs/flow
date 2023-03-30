# flow

This project models prospective material flows of steel and concrete in EU buildings as well as selected circular economy actions and bundles based on the [Open Dynamic Material Systems Model (ODYM)](https://github.com/indecol/odym). The status presented here maps to newTRENDs Deliverable 6.1.

## Overview
The general structure of the model is shown below. Within the newTRENDs-project the model was used to link the building stock model Invert/EE-Lab and the industry model FORECAST-Industry. However, the structure also allows use with other models as long as the data structure of the building stock and production processes matches.
![image](https://user-images.githubusercontent.com/96481739/228776417-817e0f5b-d995-46d1-b239-28d64ba7c2c5.png)
From Invert/EE-Lab, the building stock development was used. These data cannot be published and are only represented by dummy data in this project. For the use of the model, appropriate data must be inserted in the respective files. An overview of the required data structure is shown below.
![image](https://user-images.githubusercontent.com/96481739/228778432-d9c0b71b-c60c-4d9f-b74e-a9d5b506475c.png)
The modelled material flows were then fed into FORECAST-Industry. This step cannot be shown here.

## Requirements
The requirements for the model is stated in the file "requirements.txt". To import ODYM directly as a package, the directory must be installed (pip install .)

## Structure
The input data for the model can be found under "docs" in the model folder. The classification of flows and stocks can be found in the classification file and the configuration of the model can be found in the config file in the same folder. The results are available under "results" in the model folder. The underlying calculation is located under "src" in the model folder.

## Publications and further information
More information on the models will be available here:
- Paper on the building model (coming soon)
- newTRENDs D6.1 (coming soon)

If you want to use flow for your research, we would appreciate it is you would cite both mentioned sources.

## Contact
meta.thurid.lotz@isi.fraunhofer.de
