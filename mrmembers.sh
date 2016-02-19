#!/bin/bash
ZIP=$1
unzip $ZIP
CSV=`echo ${ZIP%.zip}".csv"`
echo $CSV
#~/bin/mrmembers.py $CSV
