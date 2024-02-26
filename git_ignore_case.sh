#!/bin/bash

set -e

#forms, classes and modules
for file in $(git status --porcelain | grep -E "^.{1}M" | grep -E -v "^R" | cut -c 4-| grep  -e "\.frm" -e "\.bas" -e "\.cls" -e "\.Dsr"); do
	ORIGFILE=$(mktemp)
	PATCHFILE=$(mktemp)
	git cat-file -p :$file > $ORIGFILE
	diff -i --strip-trailing-cr $ORIGFILE $file > $PATCHFILE || true
	patch -s $ORIGFILE < $PATCHFILE
	cp  $ORIGFILE $file
	rm  $ORIGFILE $PATCHFILE 
	unix2dos --quiet $file
done

#projects
for file in $(git status --porcelain | cut -c 4-| grep -e "\.vbp$"); do
	echo $file
	ORIGFILE=$(mktemp)
	PATCHFILE=$(mktemp)
	PROCESSEDFILE=$(mktemp)
	echo $file $ORIGFILE $PATCHFILE $PROCESSEDFILE
	git cat-file -p :$file > $ORIGFILE
	cat $ORIGFILE | cut -d "#" -f 1-3 > $PROCESSEDFILE
	diff -i --strip-trailing-cr $PROCESSEDFILE <(cat $file| cut -d "#" -f 1-3) > $PATCHFILE || true
	patch -s $ORIGFILE < $PATCHFILE
	cp $ORIGFILE $file
	unix2dos --quiet $file
done