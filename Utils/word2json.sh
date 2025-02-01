#!/bin/bash
# /root/file/1-50/36.docx


dataset_dir=/root/data
n_total=50
ratio=correct

for i in $(seq 36 50)
do
    input_word="/root/file/1-50/${i}.docx"
    echo "$input_word is in process"
    python /root/Med-NLFT/Utils/word2json.py \
        --input_word "${input_word}" \
        --dataset_dir "${dataset_dir}" \
        --n_total "${n_total}" \
        --ratio "${ratio}"
done

