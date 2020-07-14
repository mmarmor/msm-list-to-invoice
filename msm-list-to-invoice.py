r"""Convert Marmor Dynalist Time-Tracking to Invoice Format
Michael Marmor | www.michaelmarmor.com | Created on 4-July-2020
msm-list-to-invoice.py

This script takes OPML formatted files exported from Dynalist and converts
them into the format needed for Michael's invoicing system.

Input is in OPML, but the content is an outline that looks like this:

July 2020
    1-Jul-2020
        Task one name (3)
        Task two name  (1)
        Follow-up phone call (.25)
        Call that other guy (2.5)
    2-Jul-2020
        Call with Tom to debrief and plan (.5)
    3-Jul-2020
        Introduction and Vendor Demo (1)
        Call with Ted (1.75)

Output created by the script looks like this:

1-Jul-2020<tab>Task one name<tab>3
1-Jul-2020  Task two name   1
1-Jul-2020  Follow-up phone call    .25
1-Jul-2020  Call that other guy 2.5
2-Jul-2020  Call with Tom to debrief and plan   .5
3-Jul-2020  Introduction and Vendor Demo    1
3-Jul-2020  Call with Ted   1.75

Usage:

$ python msm-list-to-invoice.py filename.opml
Will write filename.txt to the same directory

I'm freezing the code with pyinstaller to run on Windows without
dependencies, so will end up calling it like this:

C:\Users\Michael> msm-list-to-invoice.exe filename.opml


Useful Resources:

Script is based on this code by Hugh Wang as a starting point:
https://gist.github.com/hghwng/324cc28b007a8f650ce3aac5df099ef8

https://www.crummy.com/software/BeautifulSoup/

https://pythex.org/

https://www.pyinstaller.org/

I ran into a Win10 execution policy problem in the way VS Code switched virtual
python environments. You can get around this by running this PowerShell script
with an elevated account:

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process

For more about this see:

https://code.visualstudio.com/docs/python/python-tutorial#setup-articles

TODO:
If the regex fails to match, we give up and stuff the line into the output
for hand editing. In this case, the hours are set to !!! to get my attention. This
could be much more sophisticated. The most important thing is to not lose data!
"""

import bs4
import re
import sys

# regex to create two capture groupss, one for the task and one for the time
p = re.compile(r'(.*)\((.+)\)$')

def convert_element(lines, level=1, mydate=''):
    result = ''
    for line in lines:
        if not isinstance(line, bs4.element.Tag) or \
           line.name != 'outline':
            continue
        if level == 2:
            mydate = line.attrs.get('text', '')
        elif level == 3:
            # compiled regex p is global to avoid re-compiling in this loop
            m = p.match(line.attrs.get('text', ''))
            if m:
                hours = m.group(2).strip()
                task = m.group(1).strip()
                result += mydate + '\t' + task + '\t' + hours + '\n'
            else:
                result += mydate + '\t' + line.attrs.get('text', '').strip() + '\t!!!' + '\n'
        result += convert_element(line.children, level + 1, mydate)
    return result

def convert_file(path):
    root = bs4.BeautifulSoup(open(path, encoding='utf8'), "lxml")
    return convert_element(root.select('html body opml')[0])

def main():
    output_path = sys.argv[1][:-4] + 'txt'
    result = convert_file(sys.argv[1])
    open(output_path, 'w').write(result)

if __name__ == '__main__':
    main()
