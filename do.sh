#!/bin/bash

if [ x"$#" != x"2" ]; then
	echo "$0 repodir outputdir"
	exit 1
fi

repodir=$1; shift
outputdir=$1; shift
mkdir -p ${outputdir}

grep xref: ${repodir}/rest_api/index.adoc | sed -e 's/^.*xref:\.\///' -e 's/#.*\[/ /' -e 's/\]$//' | sort -k2,2 | while read file title; do
	# echo "=> title=${title}	file=${file}"
	# ls -l ${repodir}/rest_api/${file} | sort -rn -t ' ' -k 5,5
	output_file=${title}__$(echo ${file} | sed -e 's|/|__|g')
	echo "=> ${outputdir}/${output_file}"
	# echo "=> ${title}"
	./a.py ${repodir}/rest_api/${file} -f xlsx -o ${outputdir}/${output_file}.xlsx
done
