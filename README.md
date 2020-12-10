# Merging .xopp (Xournal++) PDF annotations
Suppose you have PDF file "Some Name.pdf" and you also have a number of annotations to this file named by the following way:
"Some Name_1.pdf.xopp", "Some Name_2.pdf.xopp", "Some Name_3.pdf.xopp" and so on.
"Some Name" is considered as a tag. For instance, this tag can be a name of a student and "Some Name.pdf" is his examination work. This script can be used to merge annotations for different "Some Name" tags in the same time.

## Ussage

python xoppmerge.py -i exam\_xopp -o xopp\_merged -p exam\_pdf -e result\_pdfs -s

exam\_xopp is the input directory with annotations,
xopp\_merged is the output directory with merged annotations in the xopp format,
exam\_pdf is the input directory with the exam pdf files
resut\_pdfs is the output directory with the merged annotations in the pdf format.

Option -s is used for score counting. You can skip this option if you doesn't need it. To use this option, each annotation should contain at least one key phrase:
Problem "problem number": "score". The example of key phrase:
Problem 3: 1

The full number of problems is 5 (hardcoded).
