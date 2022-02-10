from screenmonkey import Sequence

mySeq = Sequence()  # initializes your Sequence object
mySeq.record()  # prompts user to do actions that they wish to record
mySeq.save_excel('testSeq.xlsx')  # saves Sequence as Excel file for repeated use
