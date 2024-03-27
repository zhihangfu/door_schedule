## why i made it?
it is always a tedious work to make door & window schedules by manually searching for door tags in a plan drawing. it is the sort of work that should be handled by computers.

## what it does?
- this python script interacts with the active autocad drawing, and searches for instances of a certain block definition. then it extracts specific attributes of those instances into a table.
- it will save and open a `.csv` file as the result.

from door & window tag blocks →            |  → to table
:-------------------------:|:-------------------------:
![what to extract](https://github.com/zhihangfu/door_schedule/assets/35970192/39e5a268-f023-4bfd-9411-a9eb7f2cb742)  |  ![what you get](https://github.com/zhihangfu/door_schedule/assets/35970192/896cc078-2fa7-41bc-969d-9ebae50b11e8)


- the block instance named "门窗编号" is being searched for, and attributes named "门窗主编号1", "门窗主编号2", "门窗辅编号1", "门窗辅编号2" are extracted. however, i don't think this block definition is complying with any uiniversal standard or convention, so i made them easy to modify at the beginning part of the code.

## how to run?

1. open the autocad drawing containing the door tags (windows will recognise it as the active doc of "AutoCAD.Application");
2. run the code;
   - for the `.py` version, successful installation of [`pywin32`](https://pypi.org/project/pywin32/) package is a prerequisite;
   - for the `.exe` version, just double-click and run, dependencies are already packaged;
4. harvest your data from the automatically opened `.csv` file.
