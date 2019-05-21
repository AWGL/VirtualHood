#!/bin/bash
set -euo pipefail

seqId=<seqid>
worksheet=<worksheet>

source /home/transfer/miniconda3/bin/activate VirtualHood

for i in <path>/*/*.variables;

do referral=$(grep "referral" $i | cut -d "=" -f2);
sampleId=$(echo $i| cut -d "/" -f6);

python panCancer_report.py $seqId $sampleId $worksheet $referral

done;
source /home/transfer/miniconda/bin/deactivate


