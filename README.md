# address-auto-typer
A tool written in python3 to take pre-downloaded google sheet or excel files containing addresses, and auto typing them on formatted pages for business use of Snowpros LLC and their subcontractors.

## Purpose
I automated a repetitive and time consuming task that was to be done every day for Snow Pros LLC, making the life of crew-leaders easier and saving the company money, because usually the crew leader would either have to write out each address in the beginning of their shift and waste set-up time, or stop the truck with a full crew to write clearly(too bumpy while moving). This task was to manually write out each address from an excel file, 3 times, on a formatted piece of paper to be used for photo taking, (before, during, after) for documentation, to prove to the city of Minneapolis that we were there clearing sidewalks. Each section would be folded and and placed in the corner of a camera frame for each photo. This task takes an average of 35 seconds to do per address, and work-orders can have over 100 addresses in them. The cost savings obviously scale with more addresses per work order. Below is the format for each address. (Ignore the 80/100, I'm messy)

![Address Format Scan-1](https://user-images.githubusercontent.com/94456069/156400012-a6de49b1-33c9-41c1-aabf-68f0388bae68.png)

## Installation for Programmers using Pycharm
* copy the source code
## 1. Add Dependencies 
* altgraph	0.17.2	0.17.2
* et-xmlfile	1.1.0	1.1.0
* future	0.18.2	0.18.2
* keyboard	0.13.5	0.13.5
* lxml	4.7.1	4.8.0
* numpy	1.22.2	1.22.2
* openpyxl	3.0.9	3.0.9
* pandas	1.4.1	1.4.1
* pefile	2021.9.3	2021.9.3
* [pip](https://www.geeksforgeeks.org/how-to-install-pip-on-windows/)	22.0.3	22.0.3
* pyinstaller	4.9	4.9
* pyinstaller-hooks-contrib	2022.2	2022.2
* python-dateutil	2.8.2	2.8.2
* python-docx	0.8.11	0.8.11
* pytz	2021.3	2021.3
* pywin32-ctypes	0.2.0	0.2.0
* setuptools	60.9.2	60.9.3
* six	1.16.0	1.16.0
* termcolor	1.1.0	1.1.0

## 2. To make Executable after installing all dependencies (.exe)
* After you've copied the code in Pycharm, Go to "terminal tab" -> "local" of Pycharm and type 
`pyinstaller main.py --onefile`
* look to the upper left hand corner for "build" find main.exe, and copy paste that to wherever you wish. 

## Testing
* Download the provided Excel file 
* Make sure to specify correct paths on lines 55 and 67 

## Results
![Screenshot (184)](https://user-images.githubusercontent.com/94456069/156395898-30b3bddd-f151-4d5e-a2a7-540fdf3cb9df.png)
![Screenshot (183)](https://user-images.githubusercontent.com/94456069/156394105-10f932eb-7c8f-49e7-b78a-36f81a2dd6fa.png)
![Screenshot (182)](https://user-images.githubusercontent.com/94456069/156394142-63fcff1c-188d-452a-8f3d-97070459b477.png)

* Now you may easily print from Microsoft Word. No sharpies and no crappy handwriting required :wink:

## Notes
* You do not need to install pyinstaller or its dependencies(altgraph, future, pefile, pywin32-ctypes, pyinstaller-hooks-contrib, pyinstaller ) if you wish to just run it as a .py
* workorder.docx automatically gets overwritten every time the program is run, so there will never be multiple files taking up space on the device under use.
