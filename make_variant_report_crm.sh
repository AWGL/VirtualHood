#!/bin/bash
set -euo pipefail

seqId=<seqid>
worksheet=<worksheet_number>
path =<path>

for i in <path>/*/*.variables;

do referral=$(grep "referral" $i | cut -d "=" -f2);
sample=$(echo $i| cut -d "/" -f6);

python CRM_report.py $seqId $sample $worksheet $referral $path
done;

source /home/transfer/miniconda3/bin/deactivate
