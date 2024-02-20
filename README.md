# Overivew
This script will take a .xml file ouput by NMAP, parse the contents of the .xml, and generate a table for all hosts within a .docx file.

# Usage
Generate a .xml using the command below (can also use -oA)
```
sudo nmap -Pn -sC -sV -oX targets -iL targets.txt
```
And then generate the .docx containing the tables
```
python3 nmap2docx-ng.py -i nmap-all-ports-manual-hosts/manual-hosts.txt.xml -o nmap-all-ports-manual-hosts/manual-hosts
```
