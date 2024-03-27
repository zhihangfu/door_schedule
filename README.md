## why i made it?
it is always a tedious work to make door & window schedules by manually searching for door tags in a plan drawing. it is the sort of work that should be handled by computers.

## what it does?
- this python script interacts with the active autocad drawing, and searches for block instances with a given named. it then extracts given attributes (by name) as table columns.
it will save and open a `.csv` file as the result.

from door tag blocks →            |  → to table
:-------------------------:|:-------------------------:
![what to extract](https://github.com/zhihangfu/door_schedule/assets/35970192/39e5a268-f023-4bfd-9411-a9eb7f2cb742)  |  ![what you get](https://github.com/zhihangfu/door_schedule/assets/35970192/896cc078-2fa7-41bc-969d-9ebae50b11e8)


- the given block name is "门窗编号", and given attribute names are "门窗主编号1", "门窗主编号2", "门窗辅编号1", "门窗辅编号2". i don't think this block definition is any sort of standard, so i made them easy to change at the beginning of the code.

## how to run?

1. open the autocad drawing containing the door tags (windows will recognise it as the active doc of "AutoCAD.Application").
2. run the code.
3. harvest your data.
