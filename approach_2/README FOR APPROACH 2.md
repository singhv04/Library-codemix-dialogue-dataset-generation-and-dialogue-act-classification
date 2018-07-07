Required tools: 
-isc_tokenizer(https://bitbucket.org/iscnlp/tokenizer)
-isc_tagger(https://bitbucket.org/iscnlp/pos-tagger)
-for language detection(https://github.com/saffsd/langid.py)
-for translation(https://pypi.org/project/googletrans/)

Programming language:Python 3.6 (Anaconda:Spyder)

Libraries used:
-openpyxl (for reading the excel file)
-xlwt (for writing .xls file)

run the program in the sequence as provided as last part of their .Then copy-paste the output from spreadsheet to an excel sheet and name the excel file as provided at last in the python code that generates that spreadsheet.
Note:For convinence i have already provide all the excel sheets generated.You can move the provided excel file to another directory and run the program and compare the files created with the one's provided.
Don't delete kitni_m file as that's the input data.



for prediction run the file testing.py.

eg.
when you will run the entire code at once after importing tokenizer and langid libraries it will be like:

DataConversionWarning: Data with input dtype int32 was converted to float64 by StandardScaler.
  warnings.warn(msg, DataConversionWarning)
0.780257936508
0.130571347708
[ 0.55555556  0.875       0.85714286  0.83333333]

enter the testक्या मुझे की कितनी  किताब मिल सकती है //here we will input the dialogue and press enter


क्या मुझे की कितनी  किताब मिल सकती है
['क्या', 'मुझे', 'की', 'कितनी', 'किताब', 'मिल', 'सकती', 'है']
[('mr', -23.809372186660767), ('hi', -20.63845658302307), ('hi', -10.727448105812073), ('hi', -27.330204844474792), ('hi', -31.305294275283813), ('hi', -16.198984265327454), ('mr', -21.93232536315918), ('hi', -13.245388865470886)]
hindi word:क्या converted to english word:what
hindi word:मुझे converted to english word:me
hindi word:की converted to english word:Of
hindi word:कितनी converted to english word:How much
hindi word:किताब converted to english word:book
hindi word:मिल converted to english word:The mill
hindi word:सकती converted to english word:Can
hindi word:है converted to english word:is
what me Of How much book The mill Can is 
['what', 'me', 'Of', 'How', 'much', 'book', 'The', 'mill', 'Can', 'is']
before stemming:what
after stemming:what
before stemming:me
after stemming:me
before stemming:Of
after stemming:Of
before stemming:How
after stemming:how
before stemming:much
after stemming:much
before stemming:book
after stemming:book
before stemming:The
after stemming:the
before stemming:mill
after stemming:mill
before stemming:Can
after stemming:can
before stemming:is
after stemming:is
['what', 'me', 'Of', 'how', 'much', 'book', 'the', 'mill', 'can', 'is']
3
noun:book
verb:book
noun:mill
['what', 'me', 'Of', 'how', 'much', 'book', 'the', 'mill', 'can', 'is']
['what', 'me', 'Of']
['book', 'the', 'mill']
2
1
8
[[11], [16]]
[1]//this tells the classification class
क्या मुझे की कितनी  किताब मिल सकती है

