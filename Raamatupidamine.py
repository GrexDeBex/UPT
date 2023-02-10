from tkinter import *
from tkinter import ttk
import openpyxl as xls
from tkcalendar import DateEntry
from copy import copy

#================================================================================================
# LISAB SISESTATUD ANDMED EXCELISSE JA UUENDAB LOGISID

def Int_To_Col(Nr, Col):      # Korduv funktsioon võtab arvu, mis näitab exceli tabelis veeru kohta ja teisendab selle vastavale tähisele
    Indexes = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K",
    "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
    
    Nr -= 1
    Surplus = Nr % 26
    Col = Indexes[Surplus]
    if Nr // 26 != 0:
        Col = Int_To_Col(Nr // 26, Col) + Col
    return Col

def Formatting(Condition, Frame, Rows, Cols, Row_Main, Col_Main):   # Funktsioon lisab vajalikele exceli lahtritele kujunduse
    if Condition == "SameRow":    # Kasutatakse siis kui on lisatud uus veerg
        for Row in Rows:
            for Col in Cols:
                
                Cell = Frame[Int_To_Col(Col_Main, "") + str(Row)]
                NewCell = Frame[Int_To_Col(Col, "") + str(Row)]
                
                NewCell.font = copy(Cell.font)
                NewCell.border = copy(Cell.border)
                NewCell.fill = copy(Cell.fill)
                NewCell.number_format = copy(Cell.number_format)
                NewCell.protection = copy(Cell.protection)
                NewCell.alignment = copy(Cell.alignment)
                
                NewCell.value = Cell.value
                
    if Condition == "SameCol":     # Kasutatakse siis kui on lisatud uus rida
        for Row in Rows:
            for Col in Cols:
                
                Cell = Frame[Col + str(Row_Main)]
                NewCell = Frame[Col + str(Row)]
                
                NewCell.font = copy(Cell.font)
                NewCell.border = copy(Cell.border)
                NewCell.fill = copy(Cell.fill)
                NewCell.number_format = copy(Cell.number_format)
                NewCell.protection = copy(Cell.protection)
                NewCell.alignment = copy(Cell.alignment)
                
                NewCell.value = Cell.value
    
    if Condition == "Table":    # Kasutatakse siis kopeeritakse terve tabel
        x=0
        for Row in Rows:
            for Col in Cols:
                Cell = Frame[Col + str(Row)]
                NewCell = Frame[Col + str(Row_Main+x)]
                
                NewCell.font = copy(Cell.font)
                NewCell.border = copy(Cell.border)
                NewCell.fill = copy(Cell.fill)
                NewCell.number_format = copy(Cell.number_format)
                NewCell.protection = copy(Cell.protection)
                NewCell.alignment = copy(Cell.alignment)
                
                NewCell.value = Cell.value
                
            x+=1
#------------------------------------------------------------------------------------------------

def Insert_Prep():    # Lisab excelisse ja logidesse vajalikud algandmed
    global Names, Occupation, ProductPay, SalesPay, Wages, Stocks, E_StockValue, Balance, Date_Prep
    
    People = int(SB_People.get())  # Inimeste arv
    
    for x in range(People):   # Võtab lahtritest iga inimese kohta andmed 
        Names.append(L_Prep[x*6].get())
        Occupation.append(L_Prep[x*6+1].get())
        Wages.append(L_Prep[x*6+2].get())
        ProductPay.append(L_Prep[x*6+3].get())
        SalesPay.append(L_Prep[x*6+4].get())
        Stocks.append(L_Prep[x*6+5].get())
    
    Row1 = 6  # Aksiate lehel esimese tabeli algusrida
    Row2 = 18+People  # Teise tabeli algus rida
    
    for x in range(People-1):   # Lisab tabelitesse vajaliku arvu ridasid
        Exl_End.insert_rows(6)
        Exl_End.insert_rows(20+x)
        Exl_Produce.insert_rows(10)
        Exl_Sales.insert_rows(10)
        Exl_Wage.insert_rows(7)
    
    Cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"]
    for x in range(People-1):       # Lisab uutele ridadele kujunduse
        Formatting("SameCol", Exl_End, [Row1+x], Cols, People+5, None)
        Formatting("SameCol", Exl_End, [Row2+x], Cols, 2*People+17, None)
        Formatting("SameCol", Exl_Produce, [x+10], Cols, People+9, None)
        Formatting("SameCol", Exl_Sales, [x+10], Cols, People+9, None)
        Formatting("SameCol", Exl_Wage, [x+7], Cols, People+6, None)

            
    for x in range(People):    # Lisab tabelitesse vajalikud arvutusfunktsioonid
        Exl_End["E" + str(Row1+x)] = "=B{}*C{}".format(Row1+x, Row1+x)
        Exl_End["G" + str(Row1+x)] = "=F{}*C{}".format(Row1+x, Row1+x)
        Exl_End["H" + str(Row1+x)] = "=G{}+E{}".format(Row1+x, Row1+x)
        
        Exl_End["E" + str(Row2+x)] = "=B{}*C{}".format(Row2+x, Row2+x)
        Exl_End["G" + str(Row2+x)] = "=F{}*C{}".format(Row2+x, Row2+x)
        Exl_End["H" + str(Row2+x)] = "=G{}+E{}".format(Row2+x, Row2+x)
        
        Exl_Wage["E" + str(7+x)] = "=J{}".format(7+x)
        Exl_Wage["G" + str(7+x)] = "=SUM(C{}:F{})".format(7+x, 7+x)
        Exl_Wage["J" + str(7+x)] = "=H{}*I{}".format(7+x, 7+x)
    
    for Col in ("C", "E", "G", "H"): # Lisab tabelitesse vajalikud arvutusfunktsioonid
        Exl_End[Col + str(7+People)] = "=SUM({}6:{}{})".format(Col, Col, 5+People)
        Exl_End[Col + str(19+2*People)] = "=SUM({}{}:{}{})".format(Col, Row2, Col, 17+2*People)
    
    for Col in ("B", "D"): # Lisab tabelitesse vajalikud arvutusfunktsioonid
        Exl_Produce[Col + str(11+People)] = "=SUM({}10:{}{})".format(Col, Col, 9+People)
    
    for Col in ("B", "C", "D", "E", "F", "G"): # Lisab tabelitesse vajalikud arvutusfunktsioonid
        Exl_Sales[Col + str(11+People)] = "=SUM({}10:{}{})".format(Col, Col, 9+People)
        
    for Col in ("C", "E", "F", "G", "J"): # Lisab tabelitesse vajalikud arvutusfunktsioonid
        Exl_Wage[Col + str(8+People)] = "=SUM({}7:{}{})".format(Col, Col, 6+People)
    
    # Lisab tabelitesse vajalikud arvutus funktsioonid
    Exl_End["B" + str(13+People)] = "=B{}-B{}".format(11+People, 12+People)
    Exl_End["C" + str(11+People)] = "=C{}".format(7+People)
    Exl_End["C" + str(13+People)] = "=C{}".format(7+People)
    Exl_End["D" + str(11+People)] = "=B{}/C{}".format(11+People, 11+People)
    Exl_End["D" + str(13+People)] = "=B{}/C{}".format(13+People, 13+People)
    
    Exl_End["B" + str(25+2*People)] = "=B{}-B{}-B{}".format(22+2*People, 23+2*People, 24+2*People)
    
    
    for x in range(People):    # Lisab tabelitesse iga inimese andmed
        Exl_End["A" + str(Row1+x)] = Names[x]
        Exl_End["B" + str(Row1+x)] = float(E_StockValue.get())
        Exl_End["C" + str(Row1+x)] = int(Stocks[x])
        Exl_End["A" + str(Row2+x)] = Names[x]
        Exl_End["B" + str(Row2+x)] = float(E_StockValue.get())
        Exl_End["C" + str(Row2+x)] = int(Stocks[x])
        Exl_Produce["A" + str(x+10)] = Names[x]
        Exl_Sales["A" + str(x+10)] = Names[x]
        Exl_Wage["A" + str(x+7)] = Names[x]
        Exl_Wage["B" + str(x+7)] = Occupation[x]
        
    Row = 10 # Pearaamatu tabeli esimene rida
    
    Exl_Book.insert_rows(Row)     #Lisab tabelisse uue rea ja kujundab selle
    Cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"]
    Formatting("SameCol", Exl_Book, [Row], Cols, Row+1, None)
    
    Cols = ["A", "B", "C", "I", "K", "L", "N", "P", "R", "T", "U"]
    for Col in Cols:    # Kohendab arvutusfunktsioone
        Exl_Book[Col + str(Row+3)] = "=SUM({}10:{}{})".format(Col, Col, Row)
    Exl_Book["C13"] = "=A13-B13"
        
    Balance = 0  # Leitakse kogu aktsiate väärtus
    for Stock in Stocks:
        Balance += int(Stock) * float(E_StockValue.get())
        
        
    Exl_Book["A10"] = Balance   # Lisatakse pearaamatusse aktsiate tehing
    Exl_Book["C10"] = Balance
    Exl_Book["D10"] = Date_Prep.get()
    Exl_Book["E10"] = "S12"
    Exl_Book["F10"] = "Õpilastelt"
    Exl_Book["G10"] = "Aktsiad"
    Exl_Book["I10"] = Balance
    
    Log = open("log.txt", "r") #Loetakse kõik logid üles
    LogText = Log.readlines()
    Log.close()
    Log = open("log.txt", "w")
    
    Data_Value = [Names, Occupation, ProductPay, SalesPay, Wages, Stocks]
    Data_Text = ["Names,", "Occupation,", "ProductPay,", "SalesPay,", "Wages,", "Stocks,"]
    
    for x in range(6): # Kõik andmed teisendatakse logidele vastavaks
        Line = Data_Text[x]
        for Value in Data_Value[x]:
            Line += Value + ","
        Line += "\n"
        LogText[x] = Line
    
    # Kõik andmed teisendatakse logidele vastavaks
    LogText[7] = "Balance," + str(Balance) + ",\n"
    LogText[9] = "Produce," + "0,"*People + "\n"
    LogText[10] = "ProductEarnings," + "0,"*People + "\n"
    LogText[12] = "SalesEarnings," + "0,"*People + "\n"
    LogText[17] = "BalanceSheetData,0," + str(Balance) + ",0,\n"
    
    for Line in LogText: # Logid kantakse sisse
        Log.write(Line)

    Log.close()
    Exl.save("Usage.xlsx") 
    quit() # Suletakse programm, et vajalikke faile uuendada

#------------------------------------------------------------------------------------------------

def Insert_Book():  # Lisab tehingu pearaamatusse
    global TransactionNr, Balance, Profits
    
    Row = TransactionNr + 10  # Pearaamatu rida, kuhu tehing lisatakse
    
    Exl_Book.insert_rows(Row)  #Lisatakse uus rida ja kujundatakse
    Cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"]
    Formatting("SameCol", Exl_Book, [Row], Cols, Row+1, None)
    
    Cols = ["A", "B", "C", "I", "K", "L", "N", "P", "R", "T", "U"] # Kohendatakse funktsioone
    for Col in Cols:
        Exl_Book[Col + str(Row+3)] = "=SUM({}10:{}{})".format(Col, Col, Row)
        
    Exl_Book["C" + str(Row+3)] = "=A{}-B{}".format(Row+3, Row+3)

    Value = float(E_Value.get()) # Saadakse tehing väärtus
    
    if Source.get() == "Väljaminek": # Väljaminek on negatiivse väärtusega arvutamise jaoks
        Value = -Value
        
    if Source.get() == "Sissetulek": # Tehingu väärtus kantakse lahtrisse
        Exl_Book["A" + str(Row)] = Value 
    else:
        Exl_Book["B" + str(Row)] = -Value
    
    Exl_Book["C" + str(Row)] = Balance + Value # Tehingu andmed ja saldo kantakse tabelisse
    Exl_Book["D" + str(Row)] = Date_Book.get()
    Exl_Book["E" + str(Row)] = Doc.get()
    Exl_Book["F" + str(Row)] = Partner.get()
    Exl_Book["G" + str(Row)] = Desc.get()
    
    if Cat.get() == "Müügitulu":   # Vastavalt tehingu kategooriale kantakse see vastavasse lahtrisse
        Exl_Book["K" + str(Row)] = Value
        Profits[0] = str(float(Profits[0]) + Value)    # Profits hoiab vastava kategooria saldot meeles
    
    elif Cat.get() == "Tootmiskulu":
        Exl_Book["L" + str(Row)] = -Value
        Profits[1] = str(float(Profits[1]) - Value)
    
    elif Cat.get() == "Turunduskulu":
        Exl_Book["P" + str(Row)] = -Value
        Profits[3] = str(float(Profits[3]) - Value)
    
    elif Cat.get() == "Muu kulu":
        Exl_Book["R" + str(Row)] = -Value
        Profits[4] = str(float(Profits[4]) - Value)
    
    elif Cat.get() == "Laen":
        if Source.get() == "Sissetulek":
            Exl_Book["U" + str(Row)] = Value
        else:
            Exl_Book["T" + str(Row)] = -Value
        BalanceSheetData[0] = float(BalanceSheetData[0]) + Value  # Hoiab laenu väärtusi meeles
    
    TransactionNr += 1  # Hoiab meeles, mitu tehingut on pearaamatus
    Balance += Value   # Uuendab saldot
    
    Log = open("log.txt", "r") #Loeb logid
    LogText = Log.readlines()
    Log.close()
    Log = open("log.txt", "w")
    
    LogText[6] = "TransactionNr," + str(TransactionNr) + ",\n"  # Teisendab andemed logidele sobivaks
    LogText[7] = "Balance," + str(Balance) + ",\n"
    
    Line = "Profits," # Teisendab andemed logidele sobivaks
    for Value in Profits:
        Line += Value + ","
    Line += "\n"
    LogText[16] = Line
    
    Line = "BalanceSheetData," # Teisendab andemed logidele sobivaks
    for Value in BalanceSheetData:
        Line += str(Value) + ","
    Line += "\n"
    LogText[17] = Line
    
    for Line in LogText: # Lisab andmed logidesse
        Log.write(Line)
    
    Log.close()
    Exl.save("Usage.xlsx")
    RaiseFrame(F_Start)  # Avab avaakna

#------------------------------------------------------------------------------------------------

def Insert_Produce(): #Täiendab tootmisaruannet
    global ProductSheetNr, Produce, ProductEarnings
    
    People = len(Names) # Inimeste arv
    
    Col = 2 * ProductSheetNr + 4 # Veerg, kus on andmete lisamise järg
    
    Exl_Produce.insert_cols(Col) # Lisatakse uued veerud
    Exl_Produce.insert_cols(Col)
    
    Formatting("SameRow", Exl_Produce, range(5,16+People), (Col, Col+1), None, Col-1) # Kujundatakse veerud
    
    for x in range(People): # Lisatakse tabelisse iga inimese kohta andmed
        Exl_Produce[Int_To_Col(Col-2, "") + str(x+10)] = int(L_Produce[x].get())
        Exl_Produce[Int_To_Col(Col-1, "") + str(x+10)] = Date_Produce.get()
        
        Sum = "=B" + str(x+10)   # Lisab iga inimese reale funktsiooni, mis liidab kõik toodetud toodede arvud kokku
        for i in range(ProductSheetNr):
            Sum = Sum + "+{}{}".format(Int_To_Col(2*i+2, ""), x+10)
        
        Exl_Produce[Int_To_Col(Col+2, "") + str(x+10)] = Sum # Lisab funktsiooni
          
    Exl_Produce[Int_To_Col(Col-2, "") + str(People+11)] = "=SUM({}10:{}{})".format(Int_To_Col(Col-2, ""), Int_To_Col(Col-2, ""), People+9) # Lisab tabelisse summade funktsioonid
    Exl_Produce[Int_To_Col(Col+2, "") + str(People+11)] = "=SUM({}10:{}{})".format(Int_To_Col(Col+2, ""), Int_To_Col(Col+2, ""), People+9)


    for x in range(People): # Peab logides järge inimeste kogutoodangul
        Produce[x] = int(Produce[x]) + int(L_Produce[x].get())
    
    ProductSheetNr += 1 # Uuendab tootmisaruannete arvu
    
    for x in range(People): # Peab järge inimeste tuludel
        ProductEarnings[x] = str(float(ProductEarnings[x]) + int(L_Produce[x].get()) * float(ProductPay[x]))
    
    Log = open("log.txt", "r") # Loeb logid
    LogText = Log.readlines()
    Log.close()
    Log = open("log.txt", "w")
    
    LogText[8] = "ProductSheetNr," + str(ProductSheetNr) + ",\n" # Teisendab andmed logide kujule
    
    Line = "Produce,"  # Teisendab andmed logide kujule
    for Value in Produce:
        Line += str(Value) + ","
    Line += "\n"
    LogText[9] = Line
    
    Line = "ProductEarnings," # Teisendab andmed logide kujule
    for Value in ProductEarnings:
        Line += Value + ","
    Line += "\n"
    LogText[10] = Line
    
    for Line in LogText: # Teisendab andmed logide kujule
        Log.write(Line)
    
    Log.close()
    Exl.save("Usage.xlsx")
    RaiseFrame(F_Start)
    

#------------------------------------------------------------------------------------------------

def Insert_Sales():  # Lisab uue müügiaruande
    global SalesSheetNr, SalesEarnings
    
    People = len(Names) # Inimeste arv
    
    Sales = [] # Iga inimese müügiarv
    Price = [] # Iga inimese müügihind
    for x in range(People):
        Sales.append(int(L_Sales[2*x].get()))
        Price.append(float(L_Sales[2*x+1].get()))
    
    Col = 3 * SalesSheetNr + 5  # Veerg, kus on müügiaruande järg
    
    for x in range(3):
        Exl_Sales.insert_cols(Col) # Lisab uued veerud
    
    Exl_Sales.merge_cells("{}8:{}8".format(Int_To_Col(Col+3,""), Int_To_Col(Col+5,"")))  # Ühendab pealmised lahtrid pealkirjade jaoks
    
    Formatting("SameRow", Exl_Sales, range(6,15+People), (Col, Col+1, Col+2), None, Col-1) # Kujundab veerud
    
    for x in range(3,6):  # Lisab summade funktsioonid veergudele
        Exl_Sales[Int_To_Col(Col+x, "") + str(11+People)] = "=SUM({}10:{}{})".format(Int_To_Col(Col+x, ""), Int_To_Col(Col+x, ""), People+9)
    
    k = 2
    for Index in ("=B", "=C", "=D"): # Lisab igale inimesele kogu müügi arvutamise funktsioonid
        for x in range(People):
            Sum = Index + str(x+10)
            for i in range(1, SalesSheetNr+1):
                Sum = Sum + "+{}{}".format(Int_To_Col(3*i+k, ""), x+10)
            
            Exl_Sales[Int_To_Col(Col+k+1, "") + str(x+10)] = Sum
        k += 1
    

    Exl_Sales[Int_To_Col(Col-3, "") + "8"] = Date_Sales.get() # Lisab kuupäeva
    
    for x in range(People):   # Lisab andmed
        Exl_Sales[Int_To_Col(Col-3, "") + str(x+10)] = Sales[x]
        Exl_Sales[Int_To_Col(Col-2, "") + str(x+10)] = Sales[x] * Price[x]
        Exl_Sales[Int_To_Col(Col-1, "") + str(x+10)] = Sales[x] * Price[x] * float(SalesPay[x])
    
    
    
    SalesSheetNr += 1  # Uuendab müügiaruannete arvu
    
    for x in range(People): # Peab järge iga inimese müügituludel
        SalesEarnings[x] = str(round(float(SalesEarnings[x]) + Sales[x] * Price[x] * float(SalesPay[x]), 2))
    
    Log = open("log.txt", "r") # Loeb logid
    LogText = Log.readlines()
    Log.close()
    Log = open("log.txt", "w")
    
    LogText[11] = "SalesSheetNr," + str(SalesSheetNr) + ",\n" # Teisendab andmed logidele vastavaks
    
    Line = "SalesEarnings," # Teisendab andmed logidele vastavaks
    for Value in SalesEarnings:
        Line += Value + ","
    Line += "\n"
    LogText[12] = Line
    
    for Line in LogText: # Lisab andmed logidesse
        Log.write(Line)
    
    Log.close()
    Exl.save("Usage.xlsx")
    RaiseFrame(F_Start) # Avab avaakna
    
        
#------------------------------------------------------------------------------------------------

def Insert_Profit(): # Lisab uue kasumiaruande
    global ProfitSheetNr, Profits
    
    Row = (ProfitSheetNr+1)*16 + 4 # Uue kasumiaruande kõige esimene rida
    
    Formatting("Table", Exl_Profit, range(4,18), ("A", "B"), Row, None) # Kujundatakse tabel
    
    Exl_Profit["B"+ str(Row+8)] = "=SUM(B{}:B{})".format(Row+4, Row+7) # Tabelisse lisatakse funktsioonid
    Exl_Profit["B"+ str(Row+9)] = "=B{}-B{}".format(Row+2, Row+8)

    Exl_Profit["B" + str(Row-16)] = Date_Profit.get()
    Exl_Profit["B" + str(Row-14)] = float(Profits[0]) # Tabelisse lisatakse andmed
    Exl_Profit["B" + str(Row-12)] = float(Profits[1])
    Exl_Profit["B" + str(Row-11)] = float(Profits[2])
    Exl_Profit["B" + str(Row-10)] = float(Profits[3])
    Exl_Profit["B" + str(Row-9)] = float(Profits[4])

    BalanceSheetData[2] = float(Profits[0]) - float(Profits[1]) - float(Profits[2]) - float(Profits[3]) - float(Profits[4])  # Bilanssi jaoks hoitakse kasum meeles

    ProfitSheetNr += 1 # Uuendatakse kasumiaruannete arvu 
    
    Log = open("log.txt", "r") # Loeb logid
    LogText = Log.readlines()
    Log.close()
    Log = open("log.txt", "w")
    
    LogText[13] = "ProfitSheetNr," + str(ProfitSheetNr) + ",\n" # Teisendab andmed logide kujule
    
    Line = "BalanceSheetData," # Teisendab andmed logide kujule
    for Value in BalanceSheetData:
        Line += str(Value) + ","
    Line += "\n"
    LogText[17] = Line
    
    for Line in LogText: # Lisab andmed logidesse
        Log.write(Line)
    
    Log.close()
    Exl.save("Usage.xlsx")
    RaiseFrame(F_Start) # Avab avaakna

#------------------------------------------------------------------------------------------------

def Insert_Wage(): # Lisab uue palgalehe
    global WageSheetNr, Balance, TransactionNr
    
    People = len(Names) # Inimeste arv
    
    Bonus = [] # Võtab kõigi selle kuu preemiad
    for x in range(People):
        Bonus.append(float(L_Wage[x].get()))
        
    
    Row = (WageSheetNr+1)*(13+People) + 4 # Uue palgalehe kõige esimene rida
    
    Cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"] 
    Formatting("Table", Exl_Wage, range(4, 15+People), Cols, Row, None) # Kujundab palgalehe

    Exl_Wage["A" + str(Row)] = "PALGALEHT " + str(WageSheetNr+1)
    Exl_Wage["B" + str(Row-People-13)] = Date_Wage.get()

    Cols = ["C", "E", "F", "G", "J"]
    for Col in Cols: # Lisab summade funktsioonid
        Exl_Wage[Col + str(Row+People+4)] = "=SUM({}{}:{}{})".format(Col, 3+Row, Col, Row+People+2)
    
    for x in range(People): # Lisab vajalikud funktsioonid tabelisse
        Exl_Wage["E" + str(Row+x+3)] = "=J{}".format(Row+x+3)
        Exl_Wage["G" + str(Row+x+3)] = "=SUM(C{}:F{})".format(Row+x+3, Row+x+3)
        Exl_Wage["J" + str(Row+x+3)] = "=H{}*I{}".format(Row+x+3, Row+x+3)
        
        Exl_Wage["C" + str(Row-People-10+x)] = float(Wages[x]) # Lisab andmed tabelisse
        Exl_Wage["D" + str(Row-People-10+x)] = float(Bonus[x])
        Exl_Wage["F" + str(Row-People-10+x)] = float(SalesEarnings[x])
        Exl_Wage["H" + str(Row-People-10+x)] = float(Produce[x])
        Exl_Wage["I" + str(Row-People-10+x)] = float(ProductPay[x])


    WageSheetNr += 1 # Peab järge palgalehtede arvul
    
    Row = TransactionNr + 10 # Pearaamatu uue tehingu rida
    
    Exl_Book.insert_rows(Row) # Lisab uue rea ja kujundab selle
    Cols = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U"]
    Formatting("SameCol", Exl_Book, [Row], Cols, Row+1, None)
    
    Cols = ["A", "B", "C", "I", "K", "L", "N", "P", "R", "T", "U"]
    for Col in Cols: # Kohandab pearaamatu funktsioone
        Exl_Book[Col + str(Row+3)] = "=SUM({}10:{}{})".format(Col, Col, Row)
    Exl_Book["C" + str(Row+3)] = "=A{}-B{}".format(Row+3, Row+3)
    
    Pay = 0 # Arvutab kogu väljamakstava palga summa
    for x in range(People):
        Pay += float(SalesEarnings[x]) + float(ProductEarnings[x]) + float(Wages[x]) + Bonus[x]
    
    Balance -= Pay # Uuendab saldot
    
    Profits[2] = str(float(Profits[2]) + Pay) # Jätab meelde pearaamatu töötasu saldo
    
    Exl_Book["B" + str(Row)] = Pay # Lisab tehingu andmed pearaamatusse
    Exl_Book["C" + str(Row)] = Balance
    Exl_Book["D" + str(Row)] = Date_Wage.get()
    Exl_Book["E" + str(Row)] = "PL" + str(WageSheetNr)
    Exl_Book["F" + str(Row)] = "Õpilastelt"
    Exl_Book["G" + str(Row)] = "Palk"
    Exl_Book["N" + str(Row)] = Pay
    
    for x in range(6): # Nullib selle kuu tasud
        ProductEarnings[x] = "0"
        SalesEarnings[x] = "0"
    
    TransactionNr += 1 # Uuendab tehigute arvu
    
    Log = open("log.txt", "r") # Loeb Logid
    LogText = Log.readlines()
    Log.close()
    Log = open("log.txt", "w")
    
    LogText[6] = "TransactionNr," + str(TransactionNr) + ",\n" # Teisendab andmed logidele sobivaks
    LogText[7] = "Balance," + str(Balance) + ",\n"
    LogText[10] = "ProductEarnings," + "0,"*People + "\n"
    LogText[12] = "SalesEarnings," + "0,"*People + "\n"
    LogText[14] = "WageSheetNr," + str(WageSheetNr) + ",\n"
    
    for Line in LogText: # Uuendab logisid
        Log.write(Line)
    
    Log.close()
    Exl.save("Usage.xlsx")
    RaiseFrame(F_Start) # Avab avaakna
    

#------------------------------------------------------------------------------------------------

def Insert_Balance(): # Lisab uue bilanssi tabeli
    global BalanceSheetNr
    
    Row = (BalanceSheetNr+1)*15 + 8 # Uue tabeli esimene rida
    
    Formatting("Table", Exl_Balance, range(8, 21), ("A", "B", "C", "D"), Row, None) # Lisab tabelile kujunduse
    
    Exl_Balance["B" + str(Row+8)] = "=SUM(B{}:B{})".format(Row+4, Row+6) # Lisab vajalikud arvutus funktsioonid
    Exl_Balance["D" + str(Row+8)] = "=SUM(D{}:D{})".format(Row+4, Row+6)
    
    Exl_Balance["C" + str(Row-15)] = Date_Balance.get() # Lisab andmed tabelisse
    Exl_Balance["B" + str(Row-11)] = float(Balance)  
    Exl_Balance["D" + str(Row-11)] = float(BalanceSheetData[0])
    Exl_Balance["D" + str(Row-10)] = float(BalanceSheetData[1])
    Exl_Balance["D" + str(Row-9)] = float(BalanceSheetData[2])
    
    BalanceSheetNr += 1 # Peab järge bilansi tabelite arvul
    
    Log = open("log.txt", "r") # Loeb logid
    LogText = Log.readlines()
    Log.close()
    Log = open("log.txt", "w")
    
    LogText[15] = "BalanceSheetNr," + str(BalanceSheetNr) + ",\n" # Teisendab andmed logidele sobivaks
    
    for Line in LogText: # Uuendab logisid
        Log.write(Line)
    
    Log.close()
    Exl.save("Usage.xlsx")
    RaiseFrame(F_Start) # Avab avaakna

#------------------------------------------------------------------------------------------------

def Insert_End():
    
    People = len(Names) # Inimeste arv
    
    Donate = float(E_Donate.get()) # Annetuse suurus
    
    Stock_Sum = 0 # Kogu aktsiate arvu arvutamine
    for Stock in Stocks:
        Stock_Sum += float(Stock)
    
    for x in range(People): # Lisab tabelisse dividendi suuruse
        Exl_End["F" + str(6+x)] = (float(Balance) - float(BalanceSheetData[1])) / Stock_Sum 
        
    # Kannab andmed likvideerimistabelisse
    Exl_End["B" + str(11+People)] = float(Balance) - float(BalanceSheetData[1]) 
    Exl_End["B" + str(12+People)] = Donate
    
    for x in range(People): # Lisab tabelisse uuendatud dividendi suuruse
        Exl_End["F" + str(18+People+x)] = (float(Balance) - float(BalanceSheetData[1]) - Donate) / Stock_Sum
        
    Exl_End["B" + str(22+2*People)] = Balance # Kannab ülejäänud andmed excelise
    Exl_End["B" + str(23+2*People)] = Balance - Donate
    Exl_End["B" + str(24+2*People)] = Donate
    
    Exl_End["F3"] = Date_End.get()
    Exl_End["F" + str(15+People)] = Date_End.get()
    
    Exl.save("Usage.xlsx")
    RaiseFrame(F_Start) # Avab avaakna
#================================================================================================
# LOGIDE LUGEMINE

Log = open("log.txt", "r")

LogList = []
LogText = Log.readlines()

for Line in LogText: # Teeb igast logide reast järjendi, kust eemaldatakse rea pealkiri ja reavahe
    List = Line.split(",")
    del List[0]
    del List[-1]
    
    LogList.append(List)

Names = LogList[0] # Logid kantakse muutujatele
Occupation = LogList[1]
ProductPay = LogList[2]
SalesPay = LogList[3]
Wages = LogList[4]
Stocks = LogList[5]
TransactionNr = int(LogList[6][0])
Balance = float(LogList[7][0])
ProductSheetNr = int(LogList[8][0])
Produce = LogList[9]
ProductEarnings = LogList[10]
SalesSheetNr = int(LogList[11][0])
SalesEarnings = LogList[12]
ProfitSheetNr = int(LogList[13][0])
WageSheetNr = int(LogList[14][0])
BalanceSheetNr = int(LogList[15][0])
Profits = LogList[16]
BalanceSheetData = LogList[17]

Log.close()

#================================================================================================
# GUI LOOMINE

#------------------------------------------------------------------------------------------------
# JÄRJENDID ANDMETE SALVESTAMISEKS

L_Prep = [] # Salvestab kõigi firmaliikmete algandmed
L_Produce = [] # Salvestab tootmisaruande jaoks vajaliku
L_Sales = [] # Salvestab müügiaruande jaoks vajaliku
L_Wage = [] # Salvestab palga preemiad

#------------------------------------------------------------------------------------------------
# TÕSTAB SOOVITUD AKNA ESILE

def RaiseFrame(Frame): # Tõstab akna esile
    Frame.tkraise()
    
#------------------------------------------------------------------------------------------------
# LOOB TEKSTI

def CreateLabel(Frame, Text, Row, Column): 
    Label(Frame, text=Text, font="Arial 12").grid(row=Row, column=Column)
    
#------------------------------------------------------------------------------------------------
# LOOB NUPU

def CreateButton(Frame, Text, Command, Row, Column): 
    Button(Frame, text=Text, font="Arial 15 bold", bd=5, command=Command).grid(row=Row, column=Column)
    
#------------------------------------------------------------------------------------------------
# LOOB TEKSTI SISENDI

def CreateEntry(Frame, Width, Row, Column):
    entry = Entry(Frame, font="Arial 12", width=Width)
    entry.grid(row=Row, column=Column)
    
    return entry
    
#------------------------------------------------------------------------------------------------
# LOOB NUMBRI SISENDI

def CreateSpin(Frame, Row, Column):
    spinbox = Spinbox(Frame, font="Arial 12", width=5, from_=0, to=10000)
    spinbox.grid(row=Row, column=Column)
    
    return spinbox

#------------------------------------------------------------------------------------------------
# LOOB VALIKUTE SISENDI

def CreateCombo(Frame, Values, Row, Column):
    combobox = ttk.Combobox(Frame, font="Arial 12", values=Values)
    combobox.grid(row=Row, column=Column)
    
    return combobox

#------------------------------------------------------------------------------------------------
# LOOB KUUPÄEVA SISENDI

def CreateDate(Frame, Row, Column):
    date = DateEntry(Frame, font="Arial 12", date_pattern='d/m/y')
    date.grid(row=Row, column=Column)
    
    return date

#------------------------------------------------------------------------------------------------
# LOOB GUI ALUSE

Window = Tk() # Loob põhiakna ja selle parameetrid
Window.geometry("1000x500")

Window.grid_rowconfigure(0, weight=1)
Window.grid_columnconfigure(0, weight=1)  

#------------------------------------------------------------------------------------------------
# LOOB KÕIK AKNAD PEALE F_PREP

F_People = Frame(Window) # Loob kõik vajalikud aknad
F_Prep = Frame(Window)
F_Add = Frame(Window)
F_Book = Frame(Window)
F_Produce = Frame(Window)
F_Sales = Frame(Window)
F_Profit = Frame(Window)
F_Wage = Frame(Window)
F_Balance = Frame(Window)
F_End = Frame(Window)

F_Start = Frame(Window)

for frame in (F_Start, F_People, F_Prep, F_Add, F_Book, F_Produce, F_Sales, F_Profit, F_Wage, F_Balance, F_End): # Paneb akendele vajalikud parameetrid
    frame.grid(row=0, column=0, sticky='nwes')
    for x in range(11):
        frame.grid_rowconfigure(x, weight=1)
        frame.grid_columnconfigure(x, weight=1)

#------------------------------------------------------------------------------------------------
# F_PREP

global E_StockValue
global Date_Prep

def RaisePrep(): # Loob akna algandmete sisestamiseks
    global E_Stock, E_StockValue, Date_Prep
    
    People = int(SB_People.get()) # Inimeste arv
    
    for x in range(People): # Loob iga inimese jaoks oma rea ja sinna reale vajalikud lahtrid
        CreateLabel(F_Prep, "Inimene %i" %(x+1), x+1, 1) 
        
        E1 = CreateEntry(F_Prep, 20, x+1, 2)
        E2 = CreateEntry(F_Prep, 20, x+1, 3)
        E3 = CreateEntry(F_Prep, 5, x+1, 4)
        E4 = CreateEntry(F_Prep, 5, x+1, 5)
        E5 = CreateEntry(F_Prep, 5, x+1, 6)
        E6 = CreateEntry(F_Prep, 5, x+1, 7)
        
        L_Prep.extend([E1, E2, E3, E4, E5, E6]) # Jätab lahtrid meelde

        
    CreateLabel(F_Prep,"Aktsia väärtus", x+2, 1) # Loob aktsia väärtuse lahtri
    E_StockValue = CreateEntry(F_Prep, 5, x+2, 2) 
    
    CreateLabel(F_Prep,"Kuupäev", x+3, 1) # Loob kuupäeva lahtri
    Date_Prep = CreateDate(F_Prep, x+3, 2)
    
    Text = ("Inimenese nimi", "Amet", "Kuupalk", "Tootmistasu", "Müügitasu", "Omatud aktsiate arv")
    for x in range(6): # Lisab lahtritele pealkirjad
        CreateLabel(F_Prep, Text[x], 0, x+2)
    
    CreateButton(F_Prep, "Lõpeta", lambda:Insert_Prep(), 10, 8) # Lisab lõpetamis nupu
    
    RaiseFrame(F_Prep) # Toob akna esile
 
#------------------------------------------------------------------------------------------------
# F_START
# Loob avaakna
People = len(Names) # Inimeste arv

if People == 0:
    CreateButton(F_Start, "Uus Fail", lambda:RaiseFrame(F_People), 3, 5) # Loob nupu uue faili alustamiseks
else:
    CreateButton(F_Start, "Muuda Faili", lambda:RaiseFrame(F_Add), 3, 5) # Loob nupu faili muutmiseks

CreateButton(F_Start, "Välju", lambda:quit(), 7, 5)

#------------------------------------------------------------------------------------------------
# F_PEOPLE

CreateLabel(F_People, "Inimeste arv firmas:", 3, 4) # Kujundab akna, kus küsitakse algandmete jaoks inimeste arvu

SB_People = CreateSpin(F_People, 3, 5)

CreateButton(F_People, "Edasi", lambda:RaisePrep(), 6, 6)
CreateButton(F_People, "Tagasi", lambda:RaiseFrame(F_Start), 6, 3)


#------------------------------------------------------------------------------------------------
# F_ADD

CreateButton(F_Add, "Kassa muutus", lambda:RaiseFrame(F_Book), 1, 5)  # Lisatakse kõik raamatupidamisega seotud protseduurid
CreateButton(F_Add, "Tootmisaruanne", lambda:RaiseFrame(F_Produce), 2, 5)
CreateButton(F_Add, "Müügiaruanne", lambda:RaiseFrame(F_Sales), 3, 5)
CreateButton(F_Add, "Kasumiaruanne", lambda:RaiseFrame(F_Profit), 4, 5)
CreateButton(F_Add, "Palk", lambda:RaiseFrame(F_Wage), 5, 5)
CreateButton(F_Add, "Bilanss", lambda:RaiseFrame(F_Balance), 6, 5)
CreateButton(F_Add, "Firma lõpetamine ja aktsiate likvideerimine", lambda:RaiseFrame(F_End), 7, 5)

CreateButton(F_Add, "Tagasi", lambda:RaiseFrame(F_Start), 10, 4)

#------------------------------------------------------------------------------------------------
# F_BOOK

CreateLabel(F_Book, "Tehing:", 1, 4)   # Kujundatakse pearaamatusse tehingu lisamise aken
Source = CreateCombo(F_Book, ("Väljaminek", "Sissetulek"), 1, 5)

CreateLabel(F_Book, "Summa:", 2, 4)
E_Value = CreateEntry(F_Book, 5, 2, 5)

CreateLabel(F_Book, "Kuupäev", 3, 4)
Date_Book = CreateDate(F_Book, 3, 5)

CreateLabel(F_Book, "Dokumendi number", 4, 4)
Doc = CreateEntry(F_Book, 20, 4, 5)

CreateLabel(F_Book, "Kellelt/Kellele", 5, 4)
Partner = CreateEntry(F_Book, 20, 5, 5)

CreateLabel(F_Book, "Mille eest", 6, 4)
Desc = CreateEntry(F_Book, 20, 6, 5)

CreateLabel(F_Book, "Mis liiki tehing", 7, 4)
Cat = CreateCombo(F_Book, ("Müügitulu", "Tootmiskulu", "Turunduskulu", "Muu kulu", "Laen"), 7, 5)


CreateButton(F_Book, "Sisesta", lambda:Insert_Book(), 10, 7)
CreateButton(F_Book, "Tagasi", lambda:RaiseFrame(F_Add), 10, 3)

#------------------------------------------------------------------------------------------------
# F_PRODUCE
# Kujundatakse tootmisaruande lisamise aken
x = 2
for Name in Names: # Iga firmatöötaja jaoks lisatakse vajalikud lahtrid
    CreateLabel(F_Produce, Name + ":", x, 4)
    SB = CreateSpin(F_Produce, x, 5)
    
    L_Produce.append(SB)
    x+=1
    
CreateLabel(F_Produce, "Toodetud produkti arv", 1, 5)

CreateLabel(F_Produce, "Kuupäev", x+2, 4)
Date_Produce = CreateDate(F_Produce, x+2, 5)


CreateButton(F_Produce, "Sisesta", lambda:Insert_Produce(), 10, 7)
CreateButton(F_Produce, "Tagasi", lambda:RaiseFrame(F_Add), 10 ,3)

#------------------------------------------------------------------------------------------------
# F_SALES
# Kujundatakse müügiaruande lisamise aken
x = 2 
for Name in Names: # Iga firmatöötaja jaoks lisatakse vajalikud lahtrid
    CreateLabel(F_Sales, Name + ":", x, 4)
    SB = CreateSpin(F_Sales, x, 5)
    E = CreateEntry(F_Sales, 5, x, 6)
    
    L_Sales.extend([SB, E])
    x+=1
    
CreateLabel(F_Sales, "Müüdud produkti arv", 1, 5)
CreateLabel(F_Sales, "Tüki hind", 1, 6)

CreateLabel(F_Sales, "Kuupäev", x+2, 4)
Date_Sales = CreateDate(F_Sales, x+2, 5)

CreateButton(F_Sales, "Sisesta", lambda:Insert_Sales(), 10, 7)
CreateButton(F_Sales, "Tagasi", lambda:RaiseFrame(F_Add), 10, 3)

#------------------------------------------------------------------------------------------------
# F_PROFIT

CreateLabel(F_Profit, "Kuupäev", 4, 4)   # Kujundatakse kasumiaruande lisamise aken
Date_Profit = CreateDate(F_Profit, 4, 5)

CreateButton(F_Profit, "Loo", lambda:Insert_Profit(), 10, 7)
CreateButton(F_Profit, "Tagasi", lambda:RaiseFrame(F_Add), 10, 3)

#------------------------------------------------------------------------------------------------
# F_WAGE
# Kujundatakse palgalehe lisamise aken
x = 2
for Name in Names: # Iga firmatöötaja jaoks lisatakse vajalikud lahtrid
    CreateLabel(F_Wage, Name + ":", x, 4)
    E = CreateEntry(F_Wage, 5, x, 5)
    
    L_Wage.append(E)
    x+=1
    
CreateLabel(F_Wage, "Preemia", 1, 5)

CreateLabel(F_Wage, "Kuupäev", x+2, 4)
Date_Wage = CreateDate(F_Wage, x+2, 5)

CreateButton(F_Wage, "Loo", lambda:Insert_Wage(), 10, 7)
CreateButton(F_Wage, "Tagasi", lambda:RaiseFrame(F_Add), 10, 3)

#------------------------------------------------------------------------------------------------
# F_BALANCE

CreateLabel(F_Balance, "Kuupäev", 4, 4) # Kujundatakse bilanssi lisamise aken
Date_Balance = CreateDate(F_Balance, 4, 5)

CreateButton(F_Balance, "Loo", lambda:Insert_Balance(), 10, 7)
CreateButton(F_Balance, "Tagasi", lambda:RaiseFrame(F_Add), 10, 3)
#------------------------------------------------------------------------------------------------
# F_END

CreateLabel(F_End, "Kui palju kavatseb firma annetada?",3, 4) # Kujundatakse aktsiate aruande aken
E_Donate = CreateEntry(F_End, 5, 3, 5)

CreateLabel(F_End, "Kuupäev", 6, 4)
Date_End = CreateDate(F_End, 6, 5)

CreateButton(F_End, "Sisesta", lambda:Insert_End(), 10, 7)
CreateButton(F_End, "Tagasi", lambda:RaiseFrame(F_Add), 10, 3)

#================================================================================================
# TOIMETUSED EXCELI FAILIGA

Exl = xls.load_workbook("Usage.xlsx") # Avatakse excel

Exl_Book = Exl["PEARAAMAT"] # Avatakse kõik töölehed
Exl_Produce = Exl["TOOTMISARUANNE"]
Exl_Sales = Exl["MÜÜGIARUANNE"]
Exl_Profit = Exl["KASUM"]
Exl_Wage = Exl["PALGALEHT"]
Exl_Balance = Exl["BILANSS"]
Exl_End = Exl["Aktsia- ja likvideerimisaruanne"]

#================================================================================================

Window.mainloop() # Käivitatakse programm