### VirtualHood

## For the CRM pipeline:

# Installation 

Clone the VirtualHood repository using:

```
git clone https://github.com/AWGL/VirtualHood.git
```

# Requirements:

The required packages can be found in envs/VirtualHood

Additional requirements include:

* Runid
* Sampleid
* Worksheet number 
* Referral must be in variables file in the the form "referral=<referral>"
* poly and artefacts list in /home/transfer/pipelines/VirtualHood


# To run:


```
referral=(grep "referral" <variablesfile> | cut -d "=" -f2);

source /home/transfer/miniconda3/bin/activate VirtualHood

python CRM_report.py $seqId $sample $worksheet $referral $path

source /home/transfer/miniconda3/bin/deactivate

```
