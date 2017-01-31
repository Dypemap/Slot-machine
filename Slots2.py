#utf-8
import xlrd
import xlwt
import random
import statistics
SlotField=[[-1,-1,-1],[-1,-1,-1],[-1,-1,-1]]
#sheet.row_values(x)[y] , x - row(stroka), y - column (stolbec)
def FillArray(filename): # read xls and fill array of elements
    ElementNum = 10
    ElemArray=[]
    ElemPositionStart=0
    ElemPositionEnd=0
    ElemPosition={}
    Reward={}
    
    rb = xlrd.open_workbook(filename, formatting_info = True)
    sheet = rb.sheet_by_index(0)
    for i in range(ElementNum):
        for j in range(int(sheet.row_values(1+i)[2])):
            ElemArray.append(str(sheet.row_values(1+i)[0]))
        ElemPositionStart=int(sheet.row_values(1+i)[4])
        ElemPositionEnd = int(sheet.row_values(1+i)[5])
        key = str(sheet.row_values(1+i)[0])
        ElemPosition[key]=[ElemPositionStart,ElemPositionEnd]
        Reward[key]=int(sheet.row_values(1+i)[1])
    return [ElemArray,ElemPosition,Reward]


def FillFirstElemOnReel(ElemArray, ElemPos):

    # fill medium cell, than top and bottom cell
    ElemArr = ElemArray
    out=random.randint(1,len(ElemArr))-1
    result = ElemArr[out]
    del ElemArr[ElemPos[result][0]-1:ElemPos[result][1]]
    return [result,ElemArr]

# del a[2:4] ElemPos[ElemArr[out]][0]-1       

def FillRestElemOnReel(ElemArray):
    out = random.randint(1,len(ElemArray))-1
    return ElemArray[out]


def MatchesNum(Field,Reward): # Pods4et koli4estva matchey
    counter = 0
    SumReward=0
    
    if (Field[0][0]==Field[0][1])and(Field[0][1]==Field[0][2]): #1
        counter+=1
        SumReward+=Reward[Field[0][0]]
    if (Field[0][0]==Field[0][1])and(Field[0][1]==Field[1][2]): #2
        counter+=1
        SumReward+=Reward[Field[0][0]]
    if (Field[0][0]==Field[1][1])and(Field[1][1]==Field[0][2]): #3
        counter+=1
        SumReward+=Reward[Field[0][0]]
    if (Field[0][0]==Field[1][1])and(Field[1][1]==Field[1][2]): #4
        counter+=1
        SumReward+=Reward[Field[0][0]]
    if (Field[0][0]==Field[1][1])and(Field[1][1]==Field[2][2]): #5
        counter+=1
        SumReward+=Reward[Field[0][0]]

    if (Field[1][0]==Field[0][1])and(Field[0][1]==Field[0][2]): #6
        counter+=1
        SumReward+=Reward[Field[1][0]]
    if (Field[1][0]==Field[0][1])and(Field[0][1]==Field[1][2]): #7
        counter+=1
        SumReward+=Reward[Field[1][0]]
    if (Field[1][0]==Field[1][1])and(Field[1][1]==Field[0][2]): #8
        counter+=1
        SumReward+=Reward[Field[1][0]]
    if (Field[1][0]==Field[1][1])and(Field[1][1]==Field[1][2]): #9
        counter+=1
        SumReward+=Reward[Field[1][0]]
    if (Field[1][0]==Field[1][1])and(Field[1][1]==Field[2][2]): #10
        counter+=1
        SumReward+=Reward[Field[1][0]]
    if (Field[1][0]==Field[2][1])and(Field[2][1]==Field[1][2]): #11
        counter+=1
        SumReward+=Reward[Field[1][0]]
    if (Field[1][0]==Field[2][1])and(Field[2][1]==Field[2][2]): #12
        counter+=1
        SumReward+=Reward[Field[1][0]]

    if (Field[2][0]==Field[1][1])and(Field[1][1]==Field[0][2]): #13
        counter+=1
        SumReward+=Reward[Field[2][0]]
    if (Field[2][0]==Field[1][1])and(Field[1][1]==Field[1][2]): #14
        counter+=1
        SumReward+=Reward[Field[2][0]]
    if (Field[2][0]==Field[1][1])and(Field[1][1]==Field[2][2]): #15
        counter+=1
        SumReward+=Reward[Field[2][0]]
    if (Field[2][0]==Field[2][1])and(Field[2][1]==Field[1][2]): #16
        counter+=1
        SumReward+=Reward[Field[2][0]]
    if (Field[2][0]==Field[2][1])and(Field[2][1]==Field[2][2]): #17
        counter+=1
        SumReward+=Reward[Field[2][0]]
    return [counter,SumReward]

#--------------prostavlenie nagradi

    




    
    
# ------------------- IGRA ----------------
InitArray=FillArray('Slots.xls')
print InitArray
print "-------------------------------"
#ElementArray = FillArray('Slots.xls')[0]
#print ElementArray, len(ElementArray)

CapitalPool = [] # realizaciya itogovogo kapitala
for k in range(10):
    Bet = 100
    Capital = 100
    MatchCounter=0 # koli4estvo aktov nagrad za seriu rollov
    PoolOfMatchQty=[0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0] # i-iy element massiva - koli4stvo odnovremennix matchey v koli4estve i , 0-evoy element massiva - koli4estvo rollov bez matchey, 1-iy - kolvo rollov s 1 matchem i t.d.
    RewardPerRoll=0

    for j in range(10000):
        for i in range(3):
            SecondReelElem = FillFirstElemOnReel(InitArray[0][:],InitArray[1])
            SecondCell = SecondReelElem[0]
            RestElemArray=SecondReelElem[1]

            FirstCell=FillRestElemOnReel(RestElemArray[:])
            ThirdCell=FillRestElemOnReel(RestElemArray[:])
        #    print 'Lenta barabana:', FirstCell, SecondCell, ThirdCell
            SlotField[0][i]=FirstCell
            SlotField[1][i]=SecondCell
            SlotField[2][i]=ThirdCell

        RewardPerRoll = MatchesNum(SlotField,InitArray[2])[1]
        PoolOfMatchQty[ MatchesNum(SlotField,InitArray[2])[0] ] +=1
        
        Capital = Capital - Bet + RewardPerRoll
        if RewardPerRoll>0:
            MatchCounter+=1
    '''        
        print 'SlotField:'
        print SlotField[0]
        print SlotField[1]
        print SlotField[2]
        print 'sovpadeniy:', MatchesNum(SlotField,InitArray[2])[0], 'Nagrada:', RewardPerRoll, Capital
        print '         ' 
    '''
    CapitalPool.append(Capital)
    print '-----SUMMARY---------'
    print '4astota matchey:', MatchCounter, 'FinalCapital:', Capital, 'pool Sovpadeniy:', PoolOfMatchQty
print 'Kone4niy Kapital:'
print statistics.mean(CapitalPool), "+-" , statistics.stdev(CapitalPool)    

#print '-----proverka sovpadeniya'
#print MatchesNum([['a1','a2','a5'],['a2','a1','a2'],['a3','a3','a1']], InitArray[2])





