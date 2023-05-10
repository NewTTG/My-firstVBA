# My-firstVBA
This is my first VBA project

Hi, just to notice i'm french so i apologosize for every fault.

This code is intended to process data collected during drive tests to determine which mobile technologies are used and the amount of data being trafficked.

I think my code can be optimized, but I mostly have an error that I can't manage and would need help with.

The entire project is called "Post traitement  DT", if you want to get an idea of the code's real-time application, that's the one.

The module where I have a problem is Module 1. The TXT file I've called "THEONE".
I have capacity errors that regularly appear on these different lines:
wsCouvertureVoix.Cells(30, "M").value = WorksheetFunction.Round(sum1G / count1G, 2)
wsCouvertureVoix.Cells(30, "I").value = WorksheetFunction.Round((count1GNonZero / count1G) * 100, 2)
wsCouvertureVoix.Cells(31, "M").value = WorksheetFunction.Round(sum1GB / count1GB, 2)
wsCouvertureVoix.Cells(31, "I").value = WorksheetFunction.Round((count1GBNonZero / count1GB) * 100, 2)
wsCouvertureVoix.Cells(row, "M").value = WorksheetFunction.Round(sum / countNonZero, 2)
wsCouvertureVoix.Cells(row, "I").value = Round((countNonZero / count) * 100, 2)

The problem is that it tries to divide by 0, which is not possible. I need it to take the values equal to 0 to calculate an average of the throughput, but if I only have values equal to 0, then I would like it to display 1.

I'm not sure if I'm being clear; I'm available on Discord: Hug0#3360

Thank you.
