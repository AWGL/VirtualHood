# VirtualHood

## For the CRM pipeline:

### Installation

Clone the VirtualHood repository using:

```
git clone https://github.com/AWGL/VirtualHood.git
```

### Requirements:

The required packages can be found in envs/VirtualHood

Additional requirements include:

* Runid
* Sampleid
* Worksheet number
* Referral must be in variables file in the the form "referral=<referral>"
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

```
source /home/transfer/miniconda3/bin/activate VirtualHood

python panCancer_report.py <seqId> <sampleid> <worksheet> <referral>

source /home/transfer/miniconda3/bin/deactivate
```

