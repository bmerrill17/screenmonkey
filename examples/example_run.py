from screenmonkey import Sequence

mySeq = Sequence()  # initializes your Sequence object
mySeq.load_excel('testSeq.xlsx')  # loads a saved Sequence of actions
mySeq.run()  # prompts user to prepare screen, then executes actions
