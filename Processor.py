

import nltk
from nltk import word_tokenize
from nltk.model import count_ngrams #for training the model, returns an instance of the 'NgramCounter' class
from nltk.model import LaplaceNgramModel
from nltk.model import LidstoneNgramModel

from nltk.model import build_vocabulary


class Processor(object):
    
    
    def __init__(self):
        
        print('\n######################################################################################################\nRunning the Processor\n######################################################################################################\n');
        
    
    # read the content of a text file
    # @Param directory: directory of the given file
    def read_data(self, directory):
        
        with open(directory, 'r') as fileContent:
            data = fileContent.read();
            #print(data);
        return data;
    
    
if __name__ == '__main__':
    processor = Processor();
    
    lbTrain = open('LB-Train.txt').read()
    mbTrain = open('MB-train.txt').read()
    
    
    # print "\n\nLB Train is "+ str(len(lbTrain)) +" :\n\n\n"
    # print lbTrain
    # print "\n\n\n"
    
    lbTraintokens = word_tokenize(lbTrain)
    mbTraintokens = word_tokenize(mbTrain)
    
        
    # print "\n\nLB lbTraintokens "+ str(len(lbTraintokens)) +"  is:\n\n\n"
 #    print lbTraintokens
 #    print "\n\n\n"
    
    lbTrainwords = [w.lower() for w in lbTraintokens]
    mbTrainwords = [w.lower() for w in mbTraintokens]
    
    
    print "\n\nLB lbTrainwords "+ str(len(lbTrainwords)) 
    print "\n\nMB mbTrainwords "+ str(len(mbTrainwords)) 
    
#     print lbTrainwords
#     print "\n\n\n"
    
    
    lbTrainvocab = build_vocabulary(2, lbTrainwords)
    mbTrainvocab = build_vocabulary(2, mbTrainwords)
    
    # print "\n\nLB lbTrainvocab.keys "+ str(len(lbTrainvocab)) +"  is:\n\n\n"
#     print lbTrainvocab.keys()
#     print "\n\n\n the vocab is"
#     print lbTrainvocab
#     print "\n\n\n"
    
    
    LB_bigram_counts = count_ngrams(2, lbTrainvocab, lbTrainwords)
    MB_bigram_counts = count_ngrams(2, mbTrainvocab, mbTrainwords)
    
    # print "\n\nLB LB_bigram_counts is:\n\n\n"
#     print LB_bigram_counts
#     print "\n\n\n"
    
    LB_Laplace_bigram_model=LaplaceNgramModel(LB_bigram_counts)
    MB_Laplace_bigram_model=LaplaceNgramModel(MB_bigram_counts)
    
    '''
    print("Building language model...")
    est = lambda fdist, bins: LaplaceProbDist(fdist)
    LB_bigram_model = NgramModel(2, lbTrain1words, estimator=est)
    '''
    #LB_bigram_model = LaplaceNgram(LB_bigram_counts)
    
    
    
    #LB_model = NgramModel(2, lbTrain1, True, False, lambda f, b:LaplaceProbDist(f))
    
    
    lbTest = open('LB-Test.txt').read()
    mbTest = open('MB-test.txt').read()
    
    # print "\n\nLB lbTest is:\n\n\n"
#     print lbTest
#     print "\n\n\n"
        
    lbTesttokens = word_tokenize(lbTest)    
    lbTestwords = [w.lower() for w in lbTesttokens]
    
    mbTesttokens = word_tokenize(mbTest)    
    mbTestwords = [w.lower() for w in mbTesttokens]
        
    print "\n\nlbTest perplexity in LB_Laplace_bigram_model is: "
    print LB_Laplace_bigram_model.perplexity(lbTestwords)
    
    print "\n\nlbTrainwords perplexity in LB_Laplace_bigram_model is: "
    print LB_Laplace_bigram_model.perplexity(lbTrainwords)
    
    print "\n\nmbTrainwords perplexity in LB_Laplace_bigram_model is: "
    print LB_Laplace_bigram_model.perplexity(mbTrainwords)
    
    print "\n\nmbTest perplexity in MB_Laplace_bigram_model is: "
    print MB_Laplace_bigram_model.perplexity(mbTestwords)
    
    print "\n\nmbTrainwords perplexity in MB_Laplace_bigram_model is: "
    print MB_Laplace_bigram_model.perplexity(mbTrainwords)
    
    print "\n\nlbTrainwords perplexity in MB_Laplace_bigram_model is: "
    print MB_Laplace_bigram_model.perplexity(lbTrainwords)
    print "\n\n\n"
    

    
    '''
    print '\nCreate session and final scripts workbook\n'
    
    snapWorkBook = openpyxl.load_workbook(Processor.SNAP_WORKBOOK)
   
    taskSheet = snapWorkBook.get_sheet_by_name('Task')
    taskSheetNew = snapWorkBook.create_sheet(index=3, title='Sessions and scripts')    
    noOfEventRows = taskSheet.get_highest_row()

    eventId = 2
    
    sessionScriptDict = {}
    
    for row in range(2 , noOfEventRows + 1):
        if(taskSheet['J' + str(row)].value != '-'):
            print '\nScript for Session no: ' + str(taskSheet['D' + str(row)].value) +'\n'
            sessionScriptDict[taskSheet['D' + str(row)].value] = taskSheet['J' + str(row)].value
    
    for i in sorted(sessionScriptDict):
        taskSheetNew['A'+str(eventId)].value = i
        taskSheetNew['B'+str(eventId)].value = sessionScriptDict[i]
        eventId += 1 

    snapWorkBook.save('Sessions and Scripts.xlsx')
    '''
    
    '''
    
    print '\nStylize chats in event sequence inorder workbook\n'
    eventsWorkBook = openpyxl.load_workbook(Processor.EVENTS_INORDER_WORKBOOK)
    sheet = eventsWorkBook.get_active_sheet()
    
    fontObj1 = Font(name='Times New Roman', bold=True, size=25, italic=True)
    questionStyle = Style(font=fontObj1)
    
    fontObj2 = Font(name='Calibri',size=24)
    chatStyle = Style(font=fontObj2)
    
    for cellObj in sheet.columns[4]:
        print cellObj.value
        print '\n\n'
        if (cellObj.value == 'CHAT') :
            print '\nChat is;\n'
            print sheet.cell(row=cellObj.row,column=8).value
            print '\n\n'
            
            if (isinstance(sheet.cell(row=cellObj.row,column=8).value, basestring)):
                if ('?' in sheet.cell(row=cellObj.row,column=8).value) :
                    sheet.cell(row=cellObj.row,column=8).style = questionStyle
                else:
                    sheet.cell(row=cellObj.row,column=8).style = chatStyle
            else:
                sheet.cell(row=cellObj.row,column=8).style = chatStyle
            
    
    eventsWorkBook.save('styled chats in events.xlsx')
    '''
    
    
    '''
    print '\nCreate event sequence inorder workbook\n'
    
    snapWorkBook = openpyxl.load_workbook(Processor.SNAP_WORKBOOK)
   
    taskSheet = snapWorkBook.get_sheet_by_name('Task')
    taskSheetNew = snapWorkBook.create_sheet(index=3, title='Events in sequence')    
    noOfEventRows = taskSheet.get_highest_row()

    eventId = 1
    
    for row in range(2 , noOfEventRows + 1):
        taskSheetNew['A'+str(eventId)].value = taskSheet['A' + str(row)].value
        taskSheetNew['B'+str(eventId)].value = taskSheet['C' + str(row)].value
        taskSheetNew['C'+str(eventId)].value = taskSheet['D' + str(row)].value
        taskSheetNew['D'+str(eventId)].value = taskSheet['F' + str(row)].value
        taskSheetNew['E'+str(eventId)].value = taskSheet['G' + str(row)].value
        taskSheetNew['F'+str(eventId)].value = str(taskSheet['D' + str(row)].value) + '.' + str(eventId)
        taskSheetNew['G'+str(eventId)].value = taskSheet['H' + str(row)].value
        taskSheetNew['H'+str(eventId)].value = taskSheet['I' + str(row)].value
        eventId += 1
        
    fontObj1 = Font(name='Times New Roman', bold=True, size=24, italic=True)
    questionStyle = Style(font=fontObj1)
    
    fontObj2 = Font(name='Times New Roman',size=20)
    chatStyle = Style(font=fontObj2)
    
    for cellObj in taskSheet.columns[4]:
        if (isinstance(cellObj.value, basestring)):
            if (cellObj.value == 'CHAT') :
                if (isinstance(taskSheetNew.cell(row=cellObj.row,column=8).value, basestring)):
                    if ('?' in taskSheetNew.cell(row=cellObj.row,column=8).value) :
                        taskSheetNew.cell(row=cellObj.row,column=8).style = questionStyle
                    else:
                        taskSheetNew.cell(row=cellObj.row,column=8).style = chatStyle
                
    
    snapWorkBook.save('events in sequence.xlsx')
    '''
    
    
    
    
    '''
    processor = Processor();
    
    utterancesWorkBook = openpyxl.load_workbook(Processor.UTTERANCESINORDER_WORKBOOK)
    sheet = utterancesWorkBook.get_active_sheet()
    
    fontObj1 = Font(name='Times New Roman', bold=True, size=24, italic=True)
    styleObj1 = Style(font=fontObj1)
    
    for cellObj in sheet.columns[3]:
        if (isinstance(cellObj.value, basestring)):
            if ('?' in cellObj.value) :
                cellObj.style = styleObj1
            
    
    utterancesWorkBook.save('styled questions in Utterances.xlsx')
    '''

    '''
    print '\nCreate chat inorder workbook\n'
    
    snapWorkBook = openpyxl.load_workbook(Processor.SNAP_WORKBOOK)
   
    taskSheet = snapWorkBook.get_sheet_by_name('Task')
    taskSheetNew = snapWorkBook.create_sheet(index=3, title='Chats in sequence')    
    noOfEventRows = taskSheet.get_highest_row()

    utteranceId = 1
    
    for row in range(2 , noOfEventRows + 1):
        if (taskSheet['G' + str(row)].value == 'CHAT'):
            taskSheetNew['A'+str(utteranceId)].value = taskSheet['D' + str(row)].value
            taskSheetNew['B'+str(utteranceId)].value = taskSheet['F' + str(row)].value
            taskSheetNew['C'+str(utteranceId)].value = 'UTT' + str(utteranceId)
            taskSheetNew['D'+str(utteranceId)].value = taskSheet['I' + str(row)].value
            utteranceId += 1
    
    snapWorkBook.save('utterances in sequence.xlsx')
    '''
    
    '''
    utterancesWorkBook = openpyxl.load_workbook(Processor.UTTERANCES_WORKBOOK)
    sheet = utterancesWorkBook.get_active_sheet()
    
    fontObj1 = Font(name='Times New Roman', bold=True, size=24, italic=True)
    styleObj1 = Style(font=fontObj1)
    
    for cellObj in sheet.columns[2]:
        if (isinstance(cellObj.value, basestring)):
            if ('?' in cellObj.value) :
                cellObj.style = styleObj1
            
    
    utterancesWorkBook.save('coloredUtterances.xlsx')
    '''
    
    '''
    snapQuestionsWorkBook = openpyxl.load_workbook(Processor.QUESTIONS_WORKBOOK)
    questionSheet = snapQuestionsWorkBook.get_sheet_by_name('Questions')
    
    howQs = 0;
    
    for cellObj in questionSheet.columns[1]:
        if ('How' in cellObj.value or 'HOW' in cellObj.value or 'how' in cellObj.value):
            print cellObj.value
            howQs += 1;
            
    print '\nThe number of How questions are:\n'
    print howQs 
    print '\n\n'
    '''
    
    '''
    #make gender corresspond to questions
    
    print '\nmake gender corresspond to questions\n'
    
    snapQuestionsWorkBook = openpyxl.load_workbook(Processor.QUESTIONS_WORKBOOK)
    questionSheet = snapQuestionsWorkBook.get_sheet_by_name('Questions')
    
    for cellObj in questionSheet.columns[0]:
        if (cellObj.row != 1):
            print '\nthe value in identifier column is: ' + str(cellObj.value) + ", for row= " + str(cellObj.row) + "\n"
            questionSheet.cell(row=cellObj.row,column=3).value = snapGenderMap.genderMap[cellObj.value]
        
    snapQuestionsWorkBook.save('updatedQuestionUtterances.xlsx')
    '''
    
    '''
    #get participant gender info
    
    print '\nget participant gender info\n'
    
    snapWorkBook = openpyxl.load_workbook(Processor.SNAP_WORKBOOK)
   
    taskSheet = snapWorkBook.get_sheet_by_name('Data Summary')
    noOfEventRows = taskSheet.get_highest_row()
    
    # dictionary which stores the id and gender
    sessionData = {}
    
    for row in range(3 , noOfEventRows + 1):
        sessionId  = taskSheet['A' + str(row)].value
        gender = taskSheet['BB' + str(row)].value
        sessionData[sessionId] = gender
        
        
    #put the data into a python file so that it can be imported later if required
    print('Writing results...')
    snapStatsResultFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapGenderMap.py', 'w')
    snapStatsResultFile.write('genderMap = ' + pprint.pformat(sessionData))
    snapStatsResultFile.close()
    '''
    
    
    
    '''
    #get stats for questions 
    print '\n\nget question utterances stats'
    
    snapQuestionsWorkBook = openpyxl.load_workbook(Processor.QUESTIONS_WORKBOOK)
    questionSheet = snapQuestionsWorkBook.get_sheet_by_name('questions Asked')
    questionSheetNew = snapQuestionsWorkBook.create_sheet(index=3, title='questions Asked stats')    

    rowNumberInFinalQuestionsSheet=2;
        
    for cellObj in questionSheet.columns[1]:
        if(cellObj.value != None):
            print '\nthe value in number column is: ' + str(cellObj.value) + ", for row= " + str(cellObj.row) + "\n"
            questionSheetNew.cell(row=rowNumberInFinalQuestionsSheet,column=2).value = questionSheet.cell(row=cellObj.row,column=2).value
            questionSheetNew.cell(row=rowNumberInFinalQuestionsSheet,column=1).value = questionSheet.cell(row=cellObj.row,column=1).value
            rowNumberInFinalQuestionsSheet += 1
        
    snapQuestionsWorkBook.save('updatedQuestionUtterances.xlsx')
    '''
    
    
    ''''
    #Clean the question utterances with ones with '?'
    
    print '\n\nget question utterances'
    
    snapQuestionsWorkBook = openpyxl.load_workbook(Processor.QUESTIONS_WORKBOOK)
    questionSheet = snapQuestionsWorkBook.get_active_sheet()
    noOfRows = questionSheet.get_highest_row()
    
    
    questionsWorkBook = openpyxl.Workbook()
    finalQuestionsSheet = questionsWorkBook.get_active_sheet()
    finalQuestionsSheet.title = "Questions"
    rowNumberInFinalQuestionsSheet=2;
        
    for cellObj in questionSheet.columns[3]:
        if(cellObj.value == 1):
            print '\nthe value in ? column is: ' + str(cellObj.value) + ", for row= " + str(cellObj.row) + ", The question is: "+questionSheet.cell(row=cellObj.row,column=3).value+" \n"
            finalQuestionsSheet.cell(row=rowNumberInFinalQuestionsSheet,column=2).value = questionSheet.cell(row=cellObj.row,column=3).value
            if (questionSheet.cell(row=cellObj.row,column=2).value == 'Driver'):
                finalQuestionsSheet.cell(row=rowNumberInFinalQuestionsSheet,column=1).value = 'DRIVER_' + str(questionSheet.cell(row=cellObj.row,column=1).value)
            else:
                finalQuestionsSheet.cell(row=rowNumberInFinalQuestionsSheet,column=1).value = 'NAV_' + str(questionSheet.cell(row=cellObj.row,column=1).value)    
            rowNumberInFinalQuestionsSheet += 1
        
    questionsWorkBook.save('updatedQuestionUtterances.xlsx')
    '''
    
    
    ''''
    # creating workbook with chats and their roles and stats
    print '\ncreating new Excel workbook for chat stats with roles\n'
    #Open and create a new snap chats stats workBook
    snapChatsStatsWorkBook = openpyxl.Workbook()
    sessionSheet = snapChatsStatsWorkBook.get_active_sheet()
    #sessionSheet.title = "Sessions' Chats Stats with roles"
    
    sessionSheet['A1'] = 'Identifier'
    sessionSheet['B1'] = 'No of chats'
    sessionSheet['C1'] = 'Average chat length'
    
    rowNumber = 2;
    
    for sessionId in sorted(snapChatsWithStats.chatsWithStats):
        
            sessionEvent = snapChatsWithStats.chatsWithStats[sessionId]
        
            print '\nsessionId is: ' + str(sessionId) + '\n'
            
            print '\nFor session: \n ' + str(sessionId) + ', row number: '+ str(rowNumber) + '\n'
            
            print '\nsessionEvent is: ' 
            print sessionEvent
            print '\n'
            
            
            sessionSheet['A' + str(rowNumber)] = 'DRIVER_'+str(sessionId)
            sessionSheet['B' + str(rowNumber)] = len(sessionEvent['driverChats'])
            sessionSheet['C' + str(rowNumber)] = sessionEvent['averageDriverChatLength']
            
            rowNumber+=1
            
            sessionSheet['A' + str(rowNumber)] = 'NAV_'+str(sessionId)
            sessionSheet['B' + str(rowNumber)] = len(sessionEvent['navChats'])
            sessionSheet['C' + str(rowNumber)] = sessionEvent['averageNavChatLength']
            
            rowNumber+=1
            
    
    snapChatsStatsWorkBook.save('chatStatsWithRoleForSnap.xlsx')
    '''
    
    
    
    '''
    #creating a workbook for the binary word vectors
    print '\ncreating new Excel workbook for the binary word vectors\n '
    #Open and create a new binary word vectors workBook
    binaryWordVectorsWorkBook = openpyxl.Workbook()
    sheet = binaryWordVectorsWorkBook.get_active_sheet()
    sheet.title = "Binary Word Vectors"
    
    sheet['A1'] = 'Session Number'
    sheet['B1'] = 'Role'
    sheet['C1'] = 'Chat Utterance'
    
    columnIndex = 4
    
    stemmedWordToColumnDict = {}
    
    for tuple in snapreducedStemmedWordFrequencyList.reducedStemmedFrequencyWordList:
        sheet.cell(row=1, column=columnIndex).value = tuple[0]
        stemmedWordToColumnDict[tuple[0]] = columnIndex
        columnIndex +=1
        
    print len(stemmedWordToColumnDict) 
    
    stemmedWordToColumnDictFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapStemmedWordToColumnDict.py', 'w')
    stemmedWordToColumnDictFile.write('reducedLemmatizedWordFrequencyList = ' + pprint.pformat(stemmedWordToColumnDict))
    stemmedWordToColumnDictFile.close()
        
    
    rowIndex = 2
    

    for sessionNumber, sessionChat in snapChatsWithStats.chatsWithStats.iteritems():
        for driverChatUtterance in sessionChat['driverChats']:
            sheet['A' + str(rowIndex)] = sessionNumber
            sheet['B' + str(rowIndex)] = 'Driver'
            sheet['C' + str(rowIndex)] = driverChatUtterance
            if type(driverChatUtterance) == types.BooleanType:
                driverChatUtterance = str(driverChatUtterance)
            driverChatWordList = nltk.word_tokenize(driverChatUtterance)
            for driverChatWord in driverChatWordList:
                stemmedDriverChatWord = stem(driverChatWord)
                if stemmedDriverChatWord in stemmedWordToColumnDict:
                    sheet.cell(row=rowIndex, column=stemmedWordToColumnDict[stemmedDriverChatWord]).value =  1
            rowIndex +=1            
        for navChatUtterance in sessionChat['navChats']:
            sheet['A' + str(rowIndex)] = sessionNumber
            sheet['B' + str(rowIndex)] = 'Navigator'
            sheet['C' + str(rowIndex)] = navChatUtterance
            if type(navChatUtterance) == types.BooleanType:
                navChatUtterance = str(navChatUtterance)
            navChatWordList = nltk.word_tokenize(navChatUtterance)
            for navChatWord in navChatWordList:
                stemmedNavChatWord = stem(navChatWord)
                if stemmedNavChatWord in stemmedWordToColumnDict:
                    sheet.cell(row=rowIndex, column=stemmedWordToColumnDict[stemmedNavChatWord]).value =  1
            rowIndex +=1
                        
    
    binaryWordVectorsWorkBook.save('binaryWordVectorsWorkBook.xlsx')    
    


    '''
    
    
    ''''
    #removing the frequency '1' words from the lemmatized word list to get the reduced list
    print('Obtaining the reduced lemmatized word list (removing words of frequency 1)...')
    reducedLemmatizedWordFrequencyList = []
    
    for tuple in snapWordFrequencyList.frequencyWordList:
        if tuple[1] != 1:
            reducedLemmatizedWordFrequencyList.append(tuple)
        
    print reducedLemmatizedWordFrequencyList  
    print('Writing the reduced lemmatized word frequency list...')
    reducedLemmatizedWordFrequencyListFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapReducedLemmatizedWordFrequencyList.py', 'w')
    reducedLemmatizedWordFrequencyListFile.write('reducedLemmatizedWordFrequencyList = ' + pprint.pformat(reducedLemmatizedWordFrequencyList))
    reducedLemmatizedWordFrequencyListFile.close()
    '''
    
    
    '''
    #removing the frequency '1' words from the stemmed word list to get the reduced list
    print('Obtaining the reduced stemmed word list (removing words of frequency 1)...')
    reducedStemmedWordFrequencyList = []
    
    for tuple in snapStemmedWordFrequencyList.frequencyWordList:
        if tuple[1] != 1:
            reducedStemmedWordFrequencyList.append(tuple)
        
    print reducedStemmedWordFrequencyList  
    print('Writing the reduced stemmed word frequency list...')
    reducedStemmedWordFrequencyListFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapreducedStemmedWordFrequencyList.py', 'w')
    reducedStemmedWordFrequencyListFile.write('reducedStemmedFrequencyWordList = ' + pprint.pformat(reducedStemmedWordFrequencyList))
    reducedStemmedWordFrequencyListFile.close()
    '''
    
    
    
    '''
    #get frequency of words in the stemmed utterances
    print('Obtaining the frequencies of words in the stemmed word list...')
    counts = Counter(snapStemmedWordList.stemmedWordList)
    
    print counts.most_common()    
    
    print('Writing the word frequency list...')
    stemmedWordFrequencyListFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapStemmedWordFrequencyList.py', 'w')
    stemmedWordFrequencyListFile.write('frequencyWordList = ' + pprint.pformat(counts.most_common()))
    stemmedWordFrequencyListFile.close()
    '''
    
    
    
    '''
    #get frequency of words in the lemmatized utterances
    print('Obtaining the frequencies of words in the lemmatized word list...')
    counts = Counter(snapLemmatizedWordList.lemmatizedWordList)
    
    print counts.most_common()    
    
    print('Writing the word frequency list...')
    wordFrequencyListFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapWordFrequencyList.py', 'w')
    wordFrequencyListFile.write('frequencyWordList = ' + pprint.pformat(counts.most_common()))
    wordFrequencyListFile.close()
    
    '''
    
    ''''
    #remove duplicates from the wordlist
    print('Removing duplicates from the word list...')
    uniqueWordList = list(set(snapLemmatizedWordList.lemmatizedWordList))
    print('Writing the unique word list...')
    uniqueWordListFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapUniqueWordList.py', 'w')
    uniqueWordListFile.write('uniqueWordList = ' + pprint.pformat(uniqueWordList))
    uniqueWordListFile.close()
    '''
    
    
    
    
    ''''
    #lemmatize the filtered word list
    
    print('Lemmatizing the word list...')
    
    nltk.download("wordnet")
    
    lemma = nltk.wordnet.WordNetLemmatizer()
    
    lemmatizedWordList = [lemma.lemmatize(word) for word in snapFilteredWordList.filteredWordList]
    
    print('Writing the lemmatized word list...')
    lemmatizedWordListFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapLemmatizedWordList.py', 'w')
    lemmatizedWordListFile.write('lemmatizedWordList = ' + pprint.pformat(lemmatizedWordList))
    lemmatizedWordListFile.close()
    '''
    
    
    
    '''
    #stem the word list. "Stemmers remove morphological affixes from words, leaving only the word stem."
    stemmedWordList = [stem(word) for word in snapFilteredWordList.filteredWordList]
    print('Writing the stemmed word list...')
    stemmedWordListFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapStemmedWordList.py', 'w')
    stemmedWordListFile.write('stemmedWordList = ' + pprint.pformat(stemmedWordList))
    stemmedWordListFile.close()
    '''
    
   
    '''
    #get filtered words with stopwords removed
    print '\n\nFiltering the chat word list to remove stop words...\n\n'
    nltk.download("stopwords")
    filtered_words = [word for word in snapWordList.wordList if word not in stopwords.words('english')]
    #put the data into a python file so that it can be imported later if required
    print('Writing the filtered word list (with stop words removed)...')
    fileteredWordListFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapFilteredWordList.py', 'w')
    fileteredWordListFile.write('filteredWordList = ' + pprint.pformat(filtered_words))
    fileteredWordListFile.close()
    '''
    
    
    ''''
    snapWorkBook = openpyxl.load_workbook(Processor.SNAP_WORKBOOK)
    '''
    ''''
    type(snapWorkBook)
    snapWorkBook.get_sheet_names()
    '''
    ''''
    taskSheet = snapWorkBook.get_sheet_by_name('Task')
    noOfEventRows = taskSheet.get_highest_row()
    '''
    #eventColumn = taskSheet.columns[6]
    
    ''''
    for cellObj in eventColumn:
            print(cellObj.value)
    '''
    ''''
    # dictionary which stores the sessionid, its events and the number of event.
    # sessionId: {EventName: No of events}
    sessionData = {}
    
    for row in range(2, noOfEventRows + 1):
        # Each row in the spreadsheet has data for one event.
        session  = taskSheet['D' + str(row)].value
        event = taskSheet['G' + str(row)].value

               # Make sure the key for this session exists.
        sessionData.setdefault(session, {})
               # Make sure the key for this event in this state exists.
        sessionData[session].setdefault(event, 0)

               # Each row represents one event, so increment by one.
        sessionData[session][event] += 1
        
        
    #put the data into a python file so that it can be imported later if required
    print('Writing results...')
    snapStatsResultFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapStats.py', 'w')
    snapStatsResultFile.write('sessionData = ' + pprint.pformat(sessionData))
    snapStatsResultFile.close()
        
        
    orderedSessionData = collections.OrderedDict(sorted(sessionData.items()))
    
    print '\ncreating new Excel workbook\n '
    #Open and create a new snap results workBook
    snapStatsWorkBook = openpyxl.Workbook()
    sessionSheet = snapStatsWorkBook.get_active_sheet()
    sessionSheet.title = 'Session Wide Stats'
    
    sessionSheet['A1'] = 'Session Number'
    sessionSheet['B1'] = 'START'
    sessionSheet['C1'] = 'CREATE'
    sessionSheet['D1'] = 'DELETE'
    sessionSheet['E1'] = 'SNAP'
    sessionSheet['F1'] = 'UNSNAP'
    sessionSheet['G1'] = 'PARAM'
    sessionSheet['H1'] = 'CAT'
    sessionSheet['I1'] = 'SPRITE'
    sessionSheet['J1'] = 'RUN'
    sessionSheet['K1'] = 'CHAT'
    sessionSheet['L1'] = 'HANGOUT'
    sessionSheet['M1'] = 'MOVE'

    rowNumber = 2;
    
    for sessionId, sessionEvent in orderedSessionData.iteritems():
        
            print '\nsessionId is: ' + sessionId + '\n'
            sessionSheet['A' + str(rowNumber)] = sessionId
            
            print '\nFor session: \n ' + str(sessionId) + ', row number: '+ str(rowNumber) + '\n'
            
            print '\nsessionEvent is: ' 
            print sessionEvent
            print '\n'
            
            
            sessionSheet['B' + str(rowNumber)] = sessionEvent['START']
            sessionSheet['C' + str(rowNumber)] = sessionEvent['CREATE']
            
            if 'DELETE' in sessionEvent:
                sessionSheet['D' + str(rowNumber)] = sessionEvent['DELETE']
                
            sessionSheet['E' + str(rowNumber)] = sessionEvent['SNAP']
            sessionSheet['F' + str(rowNumber)] = sessionEvent['UNSNAP']
            sessionSheet['G' + str(rowNumber)] = sessionEvent['PARAM']
            sessionSheet['H' + str(rowNumber)] = sessionEvent['CAT']
            
            if 'SPRITE' in sessionEvent:
             sessionSheet['I' + str(rowNumber)] = sessionEvent['SPRITE']
             
            sessionSheet['J' + str(rowNumber)] = sessionEvent['RUN']
            sessionSheet['K' + str(rowNumber)] = sessionEvent['CHAT']
            sessionSheet['L' + str(rowNumber)] = sessionEvent['HANGOUT']
            sessionSheet['M' + str(rowNumber)] = sessionEvent['MOVE']
            
            rowNumber+=1
    
    
    #snapStatsWorkBook.create_sheet(index=1, title='Inter session stats')
    
    chatStats = {}
    # {sessionID: {'driverChats': numberOfChats, 'navChats': numberOfChats} }
    for row in range(2, noOfEventRows + 1):
        # Each row in the spreadsheet has data for one event.
        session  = taskSheet['D' + str(row)].value
        event = taskSheet['G' + str(row)].value

        if event == 'CHAT':
               # Make sure the key for this session exists.
               chatStats.setdefault(session, {'driverChats': 0, 'navChats': 0})
               # Each row represents one chat event, so increment by one.
               if taskSheet['E' + str(row)].value == 'DRIVER':
                   chatStats[session]['driverChats'] += 1
               else:   
                chatStats[session]['navChats'] += 1
                   
                   
    print '\n\nThe chat stats are\n\n'               
    print chatStats        
    print '\n\n'       
    
    orderedChatData = collections.OrderedDict(sorted(chatStats.items()))
    
    snapStatsWorkBook = openpyxl.load_workbook(Processor.SNAP_STATS_WORKBOOK)
    
    sessionSheet = snapStatsWorkBook.get_active_sheet()
    
    rowNumber = 2;
    for sessionId, chatStat in orderedChatData.iteritems():
        print '\nfor sessionId: ' + str(sessionId) + ', driverChats is: '+ str(chatStat['driverChats']) + ', navChats is: '+ str(chatStat['navChats']) + ', at row: '+ str(rowNumber) + '\n'
        sessionSheet['M' + str(rowNumber)] = chatStat['driverChats']
        sessionSheet['N' + str(rowNumber)] = chatStat['navChats']
        rowNumber+=1
        
    snapStatsWorkBook.save('updatedChatStatsForSnap.xlsx')
    

    
    chats = {}
    # {sessionID: {'driverChats': [], 'navChats': []} }
    for row in range(2, noOfEventRows + 1):
        # Each row in the spreadsheet has data for one event.
        session  = taskSheet['D' + str(row)].value
        event = taskSheet['G' + str(row)].value

        if event == 'CHAT':
               # Make sure the key for this session exists.
               chats.setdefault(session, {'driverChats': [], 'navChats': []})
               # Each row represents one chat event, so increment by one.
               if taskSheet['E' + str(row)].value == 'DRIVER':
                   chats[session]['driverChats'].append(taskSheet['I' + str(row)].value)
               else:
                   chats[session]['navChats'].append(taskSheet['I' + str(row)].value)
                
                
    #put the data into a python file so that it can be imported later if required
    print('Writing chats...')
    snapChatsFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapChats.py', 'w')
    snapChatsFile.write('chats = ' + pprint.pformat(chats))
    snapChatsFile.close()
                   
                   
    print '\n\nThe chats are\n\n'               
    print chats        
    print '\n\n'
    
    '''
    
    '''
    #generate the chat stats like averages
    
    print '\n\nGenerate chat statistics...\n\n'
    
    for sessionId in sorted(snapChats.chats):
        
        sessionChats = snapChats.chats[sessionId]
        
        maxc, avg, numc, maxD, maxN, avgD, avgN, numD, numN, totalUtteranceLength, totalDriverUtteranceLength, totalNavUtteranceLength  = 0,0,0,0,0,0,0,0,0,0,0,0
        minD, minN, minc = 10000000, 10000000, 10000000
        
        driverChatsList = sessionChats['driverChats']
        navChatsList = sessionChats['navChats']
        
        print '\nthe chats for session ' + str(sessionId) + ' are;'
        
        print driverChatsList
        
        print navChatsList
        
        print '\n'
        
        sessionChats.setdefault('minimumChatLength',0)
        sessionChats.setdefault('maximumChatLength',0)
        sessionChats.setdefault('averageChatLength',0)
        sessionChats.setdefault('minimumDriverChatLength',0)
        sessionChats.setdefault('maximumDriverChatLength',0)
        sessionChats.setdefault('averageDriverChatLength',0)
        sessionChats.setdefault('minimumNavChatLength',0)
        sessionChats.setdefault('maximumNavChatLength',0)
        sessionChats.setdefault('averageNavChatLength',0)
        sessionChats.setdefault('numberOfNavChats',0)
        sessionChats.setdefault('numberOfDriverChats',0)
        sessionChats.setdefault('numberOfTotalChats',0)
        sessionChats.setdefault('totalSessionUtteranceLength',0)
        sessionChats.setdefault('totalSessionDriverUtteranceLength',0)
        sessionChats.setdefault('totalSessionNavUtteranceLength',0)
        
        for driverChatUtterance in driverChatsList:
            if type(driverChatUtterance) == types.BooleanType:
                driverChatUtterance = str(driverChatUtterance)
            driverChatUtteranceWordList = nltk.word_tokenize(driverChatUtterance)
            driverChatUtteranceLength = len(driverChatUtteranceWordList)
            if driverChatUtteranceLength < minc:
                minc = driverChatUtteranceLength
            if driverChatUtteranceLength > maxc:
                maxc = driverChatUtteranceLength
            if driverChatUtteranceLength < minD:
                minD = driverChatUtteranceLength
            if driverChatUtteranceLength > maxD:
                maxD = driverChatUtteranceLength
            totalUtteranceLength += driverChatUtteranceLength
            totalDriverUtteranceLength += driverChatUtteranceLength
            
        for navChatUtterance in navChatsList:
            if type(navChatUtterance) == types.BooleanType:
                navChatUtterance = str(navChatUtterance)
            navChatUtteranceWordList = nltk.word_tokenize(navChatUtterance)
            navChatUtteranceLength = len(navChatUtteranceWordList)
            if navChatUtteranceLength < minc:
                minc = navChatUtteranceLength
            if navChatUtteranceLength > maxc:
                maxc = navChatUtteranceLength
            if navChatUtteranceLength < minN:
                minN = navChatUtteranceLength
            if navChatUtteranceLength > maxN:
                maxN = navChatUtteranceLength
            totalUtteranceLength += navChatUtteranceLength
            totalNavUtteranceLength += navChatUtteranceLength
            


        
        
        sessionChats['minimumChatLength'] = minc
        sessionChats['maximumChatLength'] = maxc
        sessionChats['minimumDriverChatLength'] = minD
        sessionChats['maximumDriverChatLength'] = maxD
        sessionChats['numberOfDriverChats'] = len(driverChatsList)
        sessionChats['minimumNavChatLength'] = minN
        sessionChats['maximumNavChatLength'] = maxN
        sessionChats['numberOfNavChats'] = len(navChatsList)
        sessionChats['numberOfTotalChats'] = len(navChatsList)+len(driverChatsList)
        
        sessionChats['totalSessionUtteranceLength'] = totalUtteranceLength
        sessionChats['totalSessionDriverUtteranceLength'] = totalDriverUtteranceLength
        sessionChats['totalSessionNavUtteranceLength'] = totalNavUtteranceLength
        
        sessionChats['averageChatLength'] = totalUtteranceLength/(len(navChatsList)+len(driverChatsList))
        sessionChats['averageDriverChatLength'] = totalDriverUtteranceLength/len(driverChatsList)
        sessionChats['averageNavChatLength'] = totalNavUtteranceLength/len(navChatsList)  
                
                
      
    #put the data into a python file so that it can be imported later if required
    print('Writing chats again...')
    snapChatsFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapChatsWithStats.py', 'w')
    snapChatsFile.write('chatsWithStats = ' + pprint.pformat(snapChats.chats))
    snapChatsFile.close()
    
  
    '''
    
    ''''
    print '\ncreating new Excel workbook for chat stats\n '
    #Open and create a new snap chats stats workBook
    snapChatsStatsWorkBook = openpyxl.Workbook()
    sessionSheet = snapChatsStatsWorkBook.get_active_sheet()
    sessionSheet.title = "Sessions' Chats Stats"
    
    sessionSheet['A1'] = 'Session Number'
    sessionSheet['B1'] = 'max chat length'
    sessionSheet['C1'] = 'min chat length'
    sessionSheet['D1'] = 'Average chat length'
    sessionSheet['E1'] = 'Max driver chat length'
    sessionSheet['F1'] = 'min driver chat length'
    sessionSheet['G1'] = 'average driver chat length'
    sessionSheet['H1'] = 'Max nav chat length'
    sessionSheet['I1'] = 'min nav chat length'
    sessionSheet['J1'] = 'average nav chat length'
    
    rowNumber = 2;
    
    for sessionId, sessionEvent in snapChatsWithStats.chatsWithStats.iteritems():
        
            print '\nsessionId is: ' + str(sessionId) + '\n'
            sessionSheet['A' + str(rowNumber)] = sessionId
            
            print '\nFor session: \n ' + str(sessionId) + ', row number: '+ str(rowNumber) + '\n'
            
            print '\nsessionEvent is: ' 
            print sessionEvent
            print '\n'
            
            
            sessionSheet['B' + str(rowNumber)] = sessionEvent['maximumChatLength']
            sessionSheet['C' + str(rowNumber)] = sessionEvent['minimumChatLength']
            sessionSheet['D' + str(rowNumber)] = sessionEvent['averageChatLength']
                
            sessionSheet['E' + str(rowNumber)] = sessionEvent['maximumDriverChatLength']
            sessionSheet['F' + str(rowNumber)] = sessionEvent['minimumDriverChatLength']
            sessionSheet['G' + str(rowNumber)] = sessionEvent['averageDriverChatLength']
            
            sessionSheet['H' + str(rowNumber)] = sessionEvent['maximumNavChatLength']
            sessionSheet['I' + str(rowNumber)] = sessionEvent['minimumNavChatLength']
            sessionSheet['J' + str(rowNumber)] = sessionEvent['averageNavChatLength']
            
            rowNumber+=1
    
    snapChatsStatsWorkBook.save('updatedChatStatsForSnap.xlsx')
    '''
    
    
    '''
    # create a word list of all the words in the chats
    
    print '\n\nCreate the chats words list...\n\n'
    nltk.download('punkt') #nltk.download('all')
    word_list = []
    
    for sessionId, sessionEvent in snapChatsWithStats.chatsWithStats.iteritems():
        print '\nIn session :' + sessionId + ', the session event is\n'
        print sessionEvent
        print '\n\n'
        
        print '\nPrinting driver chats\n'
        
        
        for index,driverChatUtterance in enumerate(sessionEvent['driverChats']):
            if type(driverChatUtterance) == types.BooleanType:
                driverChatUtterance = str(driverChatUtterance)
            print '\n' + driverChatUtterance + '  \n'
            word_list.extend(nltk.word_tokenize(driverChatUtterance))
            print '\nthe length of word_list is '  + str(len(word_list)) + '\n'
            
        print '\nPrinting navigator chats\n'
            
        for index,navChatUtterance in enumerate(sessionEvent['navChats']):
            if type(navChatUtterance) == types.BooleanType:
                navChatUtterance = str(navChatUtterance)
            print '\n' + navChatUtterance + '  \n'
            word_list.extend(nltk.word_tokenize(navChatUtterance))
            print '\nthe length of word_list is '  + str(len(word_list)) + '\n'    
            
    #put the data into a python file so that it can be imported later if required
    print('Writing the word list...')
    wordListFile = open('/Users/mickey/Desktop/LearnDialogue/snap/snapWordList.py', 'w')
    wordListFile.write('wordList = ' + pprint.pformat(word_list))
    wordListFile.close()
    '''
    
    
    print('Done.\n')
