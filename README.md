# VirtualHood

### Installation

Clone the VirtualHood repository using:

```
git clone https://github.com/AWGL/VirtualHood.git
```
## For both CRM pipelines:

### Requirements:

The required packages can be found in envs/VirtualHood

Additional requirements include:

* Runid
* Sampleid
* Worksheet number
* Referral-must be in variables file in the the form "referral=<referral>"
* poly and artefacts list in /data/temp/artefacts_lists


## For the old CRM  pipeline:

```
source /home/transfer/miniconda3/bin/activate VirtualHood

python CRM_report.py <seqId> <sampleid> <worksheet> <referral> <NTC_folder_name>

source /home/transfer/miniconda3/bin/deactivate
```


## For the new CRM  pipeline:


```
source /home/transfer/miniconda3/bin/activate VirtualHood

python CRM_report_new_referrals.py <seqId> <sampleid> <worksheet> <referral> <NTC_folder_name>

source /home/transfer/miniconda3/bin/deactivate
```


## For the panCancer pipeline:

### Requirements:

The required packages can be found in envs/VirtualHood

Additional requirements include:

* poly and artefacts list in /data/temp/artefacts_lists

Required parameters:
* Runid
* Sampleid
* Worksheet number
* Referral- must be in variables file in the the form "referral=<referral>"

Optional:

* path - this must end with "/". If a path is not provided, the default ("/data/results/runid/RochePanCancer/") will be used.
  

```
source /home/transfer/miniconda3/bin/activate VirtualHood

python panCancer_report.py <seqId> <sampleid> <worksheet> <referral> <path>

source /home/transfer/miniconda3/bin/deactivate
```

## Tests

To run unit tests:
`python -m unittest test_panCancer_report.py`
