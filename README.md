# Merging .xopp (Xournal++) PDF annotations
Suppose you have PDF file "Some Name.pdf" and you also have a number of annotations to this file named by the following way:
"Some Name_1.pdf.xopp", "Some Name_2.pdf.xopp", "Some Name_3.pdf.xopp" and so on.
"Some Name" is considered as a tag. This tag can be a name of a student and "Some Name.pdf" is his examination work.

Run the following script:
python3 xoppmerge.py -i \<prefix to .xopp annotations> -o \<output prefix\> -p \<prefix with the original local PDFs\>
