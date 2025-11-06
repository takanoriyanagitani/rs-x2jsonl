#!/bin/sh

./x2jsonl \
	--input ./sample.xlsx \
	--sheet Sheet1 |
	jq -c
