# -*- coding: utf-8 -*-
# %% File reading
"""
Created on Wed Dec  1 15:15:30 2021

@author: Hugo Boer
"""
import pandas as pd  
import math
#import matplotlib.pyplot as plt

loc=("Vessel Data.xls")
loc2= ("Farm Data.xls")

#reading data for vessel and operation characteristics 

Data_Vessel = pd.read_excel(loc, sheet_name = 'Vessel')	     #Vessel data sheet 
Data_Operation =pd.read_excel(loc,sheet_name ='Operation')   #Operation data sheet 
Data_Farm = pd.read_excel(loc2) 		     #Farm data sheet
Data_Config = pd.read_excel(loc, sheet_name = 'Configurations')
Data_Limits = pd.read_excel(loc,sheet_name= 'Limits')


#creating dataframes for results 

T_FDN = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
WAIT_FDN = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
COST_FDN = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
T_WTB = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
WAIT_WTB = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
COST_WTB = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
T_TOT = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
COST_TOT = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])

Waitingtransfdn = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
Waitingjackfdn = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
Waitingmono = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
Waitingtp = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
Waitingposfdn = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])

Waitingtranswtb = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
Waitingposwtb = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
Waitingjackwtb = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
Waitingtow = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
Waitingnac = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])
Waitingblade = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])

WAIT_TOT = pd.DataFrame(columns =['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6', 'Option 7', 'Option 8', 'Option 9'], index=["2001","2002","2003","2004","2005","2006","2007","2008","2009","2010"])

years = [2001,2002,2003,2004,2005,2006,2007,2008,2009,2010]

                                        
for option in range (1,10):                     
    for year in years:
        
        VT_FDN = Data_Config.iloc[option,1]                #Vessel type used for foundation installation  0 = Crane barge, 1= HLV, 2 = Jack-up 
        VT_WTB = Data_Config.iloc[option,2]             #Vessel type used for turbine installation, 0= Crane barge, 1=HLV, 2= Jack-up 
                  
        
        
        startdate = str(year)+'-05-01 00:00:00'     #starting on the first day of summer on the selected year
        
        # %% Data input 
        
        #------------Farm Characteristics------ 
        
        N_FDN = Data_Farm.iloc[1,2]								#Number of foundations to be installed 
        N_WTB = Data_Farm.iloc[2,2]								#Number of turbines to be installed 
        WTB_D = Data_Farm.iloc[3,2]								#Turbine Rotor Diameter 
        INT_D = Data_Farm.iloc[4,2]								#Distance between installation locations =5*Rotor diameter 
        WD = Data_Farm.iloc[5,2]								#Water depth 
        PEN_FDN = Data_Farm.iloc[6,2]							#Penetration depth Monopile 
        N_TOW = Data_Farm.iloc[7,2]								#Number of tower pieces 
        N_BLADE = Data_Farm.iloc[8,2]							#Number of blades 
        PARK_DIS = Data_Farm.iloc[9,2]						    #Distance farm to feeding harbour 
        
        
        #--------------Vessel Characteristics----- 
        
        #Foundation vessel 
        
        TRANS_VEL_FDN= Data_Vessel.iloc[2,2+VT_FDN]               #Transition Velocity of selected Foundation installation Vessel 
        CAP_FDN= Data_Vessel.iloc[3,2+VT_FDN]					  #Loading capacity Foundation of selected Foundation installation Vessel 
        HAM_SP = Data_Vessel.iloc[5,2+VT_FDN]	                  #Hamering Speed of selected Foundation Installation Vessel 			
        AIR_FDN = Data_Vessel.iloc[7,2+VT_FDN]					  #Air gap between vessel and sea level, only for jack-up Vessels 
        PEN_LEG_FDN = Data_Vessel.iloc[8,2+VT_FDN]				  #Penetration depth of the legs of jack-up vessel 
        JACK_SP_FDN = Data_Vessel.iloc[9,2+VT_FDN]				  #Jack-up-speed, only for jack-up vessels
        CHT_C_FDN =  Data_Vessel.iloc[12,2+VT_FDN]				  #Day-rate charter cost FDN vessel 
        JACKING_FDN =Data_Vessel.iloc[14,2+VT_FDN]                #Does the vessel require jacking 
        SUPPLY_FDN = Data_Vessel.iloc[15,2+VT_FDN]
        
        
        
        #Turbine installation vessel 
        TRANS_VEL_WTB= Data_Vessel.iloc[2,2+VT_WTB]               #Transition Velocity WTB Vessel 
        CAP_WTB = Data_Vessel.iloc[4,2+VT_WTB]					  #Loading capacity WTB vessel 
        AIR_WTB = Data_Vessel.iloc[7,2+VT_WTB]					  #Air gap between vessel and sea level, only for jack-up Vessels 
        PEN_LEG_WTB =Data_Vessel.iloc[8,2+VT_WTB]				  #Penetration depth of the legs of jack-up vessel 
        JACK_SP_WTB = Data_Vessel.iloc[9,2+VT_WTB]				  #Jack-up-speed, only for jack-up vessels
        CHT_C_WTB =  Data_Vessel.iloc[12,2+VT_WTB]				  #Day-rate charter cost selected WTB vessel 
        JACKING_WTB = Data_Vessel.iloc[14,2+VT_WTB]               #Does the vessel require jacking 
        SUPPLY_WTB =Data_Vessel.iloc[15,2+VT_WTB]
        
        # #---------Operation Characteristics -----
        
        LT_FDN 		= Data_Operation.iloc[2,2+VT_FDN]			  #Loading time foundations 
        LT_WTB 		= Data_Operation.iloc[3,2+VT_WTB]			  #Loading time Turbine components (includes towers, nacelles and blades)
        POS_FDN 	= Data_Operation.iloc[4,2+VT_FDN]			  #Postitioning Time of selected Foundation Installation Vessel 
        POS_WTB 	= Data_Operation.iloc[4,2+VT_WTB]			  #Postiitioning Time of selected Wind Turbine installation vessel 
        LFT_FDN 	= Data_Operation.iloc[5,2+VT_FDN]			  #Lifting time foundation 
        LFT_TP		= Data_Operation.iloc[6,2+VT_FDN]		      #Lifting time transition piece 
        GRT_TP 		= Data_Operation.iloc[7,2+VT_FDN]			  #Grouting time transition piece 
        LFT_TOW 	= Data_Operation.iloc[8,2+VT_WTB]			  #Lifting time tower piece 
        LFT_NAC 	= Data_Operation.iloc[9,2+VT_WTB]			  #Lifting time nacelle 
        LFT_BLADE	= Data_Operation.iloc[10,2+VT_WTB]			  #lifting time blade 
        INST_BLADE 	= Data_Operation.iloc[11,2+VT_WTB]		      #installation time blade 
        
        
               #---------Operational limits ---------
       #Foundation installation 
       
        HS_TRANS_FDN = Data_Limits.iloc[2+VT_FDN,1]               #Significant wave height limit transportation Foundation Vessel 
        HS_POS_FDN = Data_Limits.iloc[2+VT_FDN,2]                 #Significant wave height limit positioning Foundation Vessel 
        HS_MONO =Data_Limits.iloc[2+VT_FDN,3]                     #Significant wave height limit monopile installation  
        TP_MONO =Data_Limits.iloc[2+VT_FDN,4]                     #Mean period limit monopile installation 
        WS_MONO = Data_Limits.iloc[2+VT_FDN,5]                    #Wind speed limit monopile installation 
        HS_TP = Data_Limits.iloc[2+VT_FDN,6]                      #Significant wave height limit transition piece installation     
        TP_TP =Data_Limits.iloc[2+VT_FDN,7]                       #Mean period limit transition piece installation 
        WS_TP =Data_Limits.iloc[2+VT_FDN,8]                       #Wind speed limit transtition piece installation 
       
       #Turbine installation
        HS_TRANS_WTB = Data_Limits.iloc[10+VT_WTB,1]              #Significant wave height limit transportation Wind turbine Vessel 
        HS_POS_WTB = Data_Limits.iloc[10+VT_WTB,2]                #Significant wave height limit positioning Wind turbine Vessel 
        HS_TOW = Data_Limits.iloc[10+VT_WTB,3]                    #Significant wave height limit tower installation 
        TP_TOW = Data_Limits.iloc[10+VT_WTB,4]                    #Mean period limit tower installation 
        WS_TOW = Data_Limits.iloc[10+VT_WTB,5]                    #Wind speed limit tower 
        HS_NAC = Data_Limits.iloc[10+VT_WTB,6]                    #Significant wave height limit nacelle installation 
        TP_NAC = Data_Limits.iloc[10+VT_WTB,7]                    #Mean period limit nacelle installation 
        WS_NAC = Data_Limits.iloc[10+VT_WTB,8]                    #Wind speed limit nacelle installation 
        HS_BLADE = Data_Limits.iloc[10+VT_WTB,9]                  #Significant wave height limit blade installation 
        TP_BLADE = Data_Limits.iloc[10+VT_WTB,10]                 #Mean period limit blade installation 
        WS_BLADE = Data_Limits.iloc[10+VT_WTB,11]                 #Wind speed limit blade installation 
        HS_JACK = 1                                               #Significant wave height limit jacking       
 
        
        #%% Weather Data input 
        df1= pd.read_table("C:/Users/hugo-/Dropbox/Research Assigment/"+str(year)+".txt") 
        df1['mean_period'] = 1/df1.peak_fr # hier kan ik gelijk de txt bestanden mee inlezen geen omvormen naar excel mer nodig 
        df1 = df1[['datetime','s_wht','wind_speed','mean_period']]
        
     
     
        
        #%% Starting parameter set up 
        
        ts= pd.to_datetime(startdate)   #creates a timestamp for the start of the year
        
        start= (ts.dayofyear-1)*24+ts.hour   #gives us the hour of the year. 
        
        
        #%% Operation steps Foundation 
        TOTAL_TIME_FDN = 0 
        stock_FDN = N_FDN 
        waiting = 0
        waitingtransfdn = 0
        waitingjackfdn = 0 
        waitingmono = 0 
        waitingtp = 0
        waitingposfdn = 0
        
        
        FDN_Installed = 0 
        FDN_Onboard = 0 
        
        
        
        if SUPPLY_FDN == 0:
            
            while FDN_Installed<N_FDN:										#Loop to run until all the desired number of turbines are installed 
                if stock_FDN > CAP_FDN:										#Check if the stock at the port is sufficent to load vessel 
                    FDN_Onboard = CAP_FDN						            #Number of foundations on board is loaded untill vessel capacity is reached 
                else:
                    FDN_Onboard = stock_FDN									#If stock level is too low load remaining foundations 
                stock_FDN = stock_FDN - FDN_Onboard							#New stocklevel after loading 
                
                 
                DUR_LOAD_FDN = FDN_Onboard*LT_FDN 							#Step 1 Loading turbines in port 
                
                TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_LOAD_FDN
                #print("Loading complete")
               # print ("stockfdn",stock_FDN)
                #print("fdn on deck after port",FDN_Onboard)
                #print("Left port")
                #print("total time after loading",TOTAL_TIME_FDN)
                DUR_TRANS_FDN = (PARK_DIS/TRANS_VEL_FDN)			#Step 2 transportation towards site 		
                #TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_TRANS_FDN
                
                
                durationtrans = math.ceil(DUR_TRANS_FDN)                         
                for i in range(0,len(df1.datetime)-start):                          #checks wether we are able to set sail with the given limits. 
                
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+duration-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                    if (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationtrans-1,['s_wht']].max().item() <= HS_TRANS_FDN): 
                        #print ('good weather window')
                        TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_TRANS_FDN
                        break
                    else: 
                       # print('bad weather window')
                        waiting = waiting+1
                        TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                        i=+1
                        waitingtransfdn = waitingtransfdn +1
                #print("total time after transit",TOTAL_TIME_FDN)
                
                while FDN_Onboard>0:
            																#Loop to run untill all foundation on board are installed 
                    DUR_POS_FDN = POS_FDN									#Step 3 positioning of vessel 
                    #TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_POS_FDN
                    #print('time after positioning', TOTAL_TIME_FDN)
                    durationpos = math.ceil(DUR_POS_FDN)                 #calculate the duration of step 2 and 3 to find the length of the weather window we need to find 
                    
                    for i in range(0,len(df1.datetime)-start):
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+durationpos-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                        if (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationpos-1,['s_wht']].max().item() <= HS_POS_FDN): 
                            #print ('good weather window')
                            TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_POS_FDN 
                            break
                        else: 
                            #print('bad weather window')
                            waiting = waiting+1
                            TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                            i=+1
                            waitingposfdn = waitingposfdn+1
                            
                            
                    if JACKING_FDN == 1:											#step 4 jacking up (only jack-up vessel)
                        DUR_JACK_FDN =(PEN_LEG_FDN + AIR_FDN + WD) /JACK_SP_FDN
                        #TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_JACK_FDN
                        durationjack = DUR_JACK_FDN 
                        
                        for i in range(0,len(df1.datetime)-start):
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+durationpos-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                            if (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationjack-1,['s_wht']].max().item() <= HS_JACK): 
                            #print ('good weather window')
                                TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_JACK_FDN
                                break
                            else: 
                            #print('bad weather window')
                                waiting = waiting+1
                                TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                                i=+1
                                waitingjackfdn = waitingjackfdn+1
                    else:
                        DUR_JACK_FDN = 0
                    
                    
                   
                    
                    #print("Time after positioning vessel:",TOTAL_TIME_FDN)
                    
                    
                    
                    DUR_HOIST_FDN = LFT_FDN 									#step 5 lifting and positioning monopile on seabed 
                    #TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_HOIST_FDN
                    #print('time after lift MP',FDN_Installed+1,':',TOTAL_TIME_FDN)
                    
                    DUR_HAM = PEN_FDN/HAM_SP          						#step 6 Hammer monopile down using hydraulic jack 
                    #TOTAL_TIME_FDN = TOTAL_TIME_FDN +DUR_HAM
                    #print('time after hammering MP',FDN_Installed+1,':',TOTAL_TIME_FDN)
                    
                    
                    durationmono= math.ceil(DUR_HOIST_FDN + DUR_HAM)
                    
                    for i in range(0,len(df1.datetime)-start):
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+durationmono-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                        if ((df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationmono-1,['s_wht']].max().item() <= HS_MONO) and (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationmono-1,['wind_speed']].max().item() <= WS_MONO) and (df1.loc[math.ceil(start+TOTAL_TIME_FDN):math.ceil(start+TOTAL_TIME_FDN+durationmono-1),['mean_period']].max().item() <= TP_MONO))  : 
                            #print ('good weather window')
                            TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_HOIST_FDN + DUR_HAM
                            break
                        else: 
                            #print('bad weather window')
                            waiting = waiting+1
                            TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                            i=+1
                            waitingmono = waitingmono+1
                    
                    DUR_HOIST_TP = LFT_TP 									#Step 7 Lifting Transition piece into place 
                    
                    #print('time after lifting TP',FDN_Installed+1, ':', TOTAL_TIME_FDN)
                    
                    DUR_INST_TP = GRT_TP									#Step 8 Installing/Grouting Transition Piece 
                    
                    #print('time after grouting TP',FDN_Installed+1,':',TOTAL_TIME_FDN)
                    
                    durationtp = math.ceil(DUR_HOIST_TP + DUR_INST_TP)
                    
                    for i in range(0,len(df1.datetime)-start):
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+durationtp-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                        if ((df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationtp-1,['s_wht']].max().item() <= HS_TP) and (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationtp-1,['wind_speed']].max().item() <= WS_TP) and (df1.loc[math.ceil(start+TOTAL_TIME_FDN):math.ceil(start+TOTAL_TIME_FDN+durationtp-1),['mean_period']].max().item() <=TP_TP))  : 
                            #print ('good weather window')
                            TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_HOIST_TP + GRT_TP
                            break
                        else: 
                            #print('bad weather window')
                            waiting = waiting+1
                            TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                            i=+1
                            waitingtp = waitingtp +1
                
                    
                    if JACKING_FDN == 1:											#step 9 jacking down (only jack-up vessel)
                        DUR_JACK_FDN =(PEN_LEG_FDN + AIR_FDN + WD) /JACK_SP_FDN
                        #TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_JACK_FDN
                        durationjack = DUR_JACK_FDN 
                        
                        for i in range(0,len(df1.datetime)-start):
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+durationpos-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                            if (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationjack-1,['s_wht']].max().item() <= HS_JACK): 
                            #print ('good weather window')
                                TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_JACK_FDN
                                break
                            else: 
                            #print('bad weather window')
                                waiting = waiting+1
                                TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                                i=+1
                                waitingjackfdn = waitingjackfdn+1
                        
                        #print('time after jacking down', TOTAL_TIME_FDN) 		
                    else:
                        DUR_JACK_FDN = 0 
                    
                    
                    
                    FDN_Installed = FDN_Installed+1
                    FDN_Onboard = FDN_Onboard -1
                    #print('time after installing foundation',FDN_Installed,':',TOTAL_TIME_FDN)
                
                    DUR_REPO_FDN = (INT_D/TRANS_VEL_FDN)                            #step 10 reposition vessel to start next installation 
                    TOTAL_TIME_FDN = TOTAL_TIME_FDN +DUR_REPO_FDN
                    #print('time after repositioning', TOTAL_TIME_FDN)
                else:
                    DUR_TRANS_FDN = (PARK_DIS/TRANS_VEL_FDN)				#Step return to port  		
                    #TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_TRANS_FDN
                    for i in range(0,len(df1.datetime)-start):                          #checks wether we are able to set sail with the given limits. 
                    
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+duration-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                        if (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationtrans-1,['s_wht']].max().item() <= HS_TRANS_FDN): 
                            #print ('good weather window')
                            TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_TRANS_FDN
                            break
                        else: 
                           # print('bad weather window')
                            waiting = waiting+1
                            waitingtransfdn= waitingtransfdn+1
                            TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                            i=+1
                    
                    #print('return to port')
            #print('TOTAL time installation foundations', math.ceil(TOTAL_TIME_FDN), 'hours')      
            T_FDN.loc[str(year),'Option '+str(option)] = math.ceil(TOTAL_TIME_FDN)
            WAIT_FDN.loc[str(year),'Option '+str(option)] = math.ceil(waiting)
            Waitingtransfdn.loc[str(year),'Option '+str(option)] = math.ceil(waitingtransfdn)
            Waitingposfdn.loc[str(year),'Option '+str(option)] = math.ceil(waitingposfdn)
            Waitingjackfdn.loc[str(year),'Option '+str(option)] = math.ceil(waitingjackfdn)
            Waitingmono.loc[str(year),'Option '+str(option)] = math.ceil(waitingmono)
            Waitingtp.loc[str(year),'Option '+str(option)] = math.ceil(waitingtp)
            
            
        elif SUPPLY_FDN == 1:
            
          #print("Left port")
          DUR_TRANS_FDN = (PARK_DIS/TRANS_VEL_FDN)			#Step 2 transportation towards site 		
          #TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_TRANS_FDN
          #print("total time after transit",TOTAL_TIME_FDN)  
          durationtrans = math.ceil(DUR_TRANS_FDN)                         
          for i in range(0,len(df1.datetime)-start):                          #checks wether we are able to set sail with the given limits. 
          
              #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+duration-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
              if (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+1+math.ceil(TOTAL_TIME_FDN)+durationtrans-1,['s_wht']].max().item() <= HS_TRANS_FDN): 
                  #print ('good weather window')
                  TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_TRANS_FDN
                  break
              else: 
                 # print('bad weather window')
                  waiting = waiting+1
                  TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                  i=+1 
                  waitingtransfdn = waitingtransfdn +1
            
          while FDN_Installed<N_FDN:										#Loop to run until all the desired number of turbines are installed 
                						
                
                    
                while FDN_Onboard>0:
            																#Loop to run untill all foundation on board are installed 
                    DUR_POS_FDN = POS_FDN									#Step 3 positioning of vessel 
                    #TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_POS_FDN
                    #print('time after positioning', TOTAL_TIME_FDN)
                    
                    if JACKING_FDN == 1:											#step 4 jacking up (only jack-up vessel)
                        DUR_JACK_FDN =(PEN_LEG_FDN + AIR_FDN + WD) /JACK_SP_FDN
                        #TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_JACK_FDN
                        #print('time after jacking up', TOTAL_TIME_FDN)			
                    else:
                        DUR_JACK_FDN = 0
                    
                    durationpos = math.ceil(DUR_POS_FDN+DUR_JACK_FDN)                 #calculate the duration of step 2 and 3 to find the length of the weather window we need to find 
                    
                    for i in range(0,len(df1.datetime)-start):
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+durationpos-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                        if (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationpos-1,['s_wht']].max().item() <=HS_POS_FDN): 
                            #print ('good weather window')
                            TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_JACK_FDN + DUR_POS_FDN 
                            break
                        else: 
                            #print('bad weather window')
                            waiting = waiting+1
                            waitingposfdn =waitingposfdn+1
                            TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                            i=+1
                    
                    
                
                    DUR_HOIST_FDN = LFT_FDN 									#step 5 lifting and positioning monopile on seabed 
                    
                   
                    DUR_HAM = PEN_FDN/HAM_SP          						#step 6 Hammer monopile down using hydraulic jack 
                    
                    
                    durationmono= math.ceil(DUR_HOIST_FDN + DUR_HAM)
                    
                    for i in range(0,len(df1.datetime)-start):
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+durationpos-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                        if ((df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationmono-1,['s_wht']].max().item() <= HS_MONO) and (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationmono-1,['wind_speed']].max().item() <= WS_MONO) and (df1.loc[math.ceil(start+TOTAL_TIME_FDN):math.ceil(start+TOTAL_TIME_FDN+durationmono-1),['mean_period']].max().item() <= TP_MONO))  : 
                            #print ('good weather window')
                            TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_HOIST_FDN + DUR_HAM
                            break
                        else: 
                            #print('bad weather window')
                            waiting = waiting+1
                            TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                            i=+1
                            waitingmono = waitingmono+1 
                    
                    DUR_HOIST_TP = LFT_TP 									#Step 7 Lifting Transition piece into place 
                    
                    
                    DUR_INST_TP = GRT_TP									#Step 8 Installing/Grouting Transition Piece 
                    
                    durationtp = math.ceil(DUR_HOIST_TP + DUR_INST_TP)
                    
                    for i in range(0,len(df1.datetime)-start):
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+durationpos-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                        if ((df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationtp-1,['s_wht']].max().item() <= HS_TP) and (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationtp-1,['wind_speed']].max().item() <= WS_TP) and (df1.loc[math.ceil(start+TOTAL_TIME_FDN):math.ceil(start+TOTAL_TIME_FDN+durationtp-1),['mean_period']].max().item() <= TP_TP))  : 
                            #print ('good weather window')
                            TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_HOIST_TP + GRT_TP
                            break
                        else: 
                            #print('bad weather window')
                            waiting = waiting+1
                            TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                            i=+1
                            waitingtp = waitingtp +1
                    
                    if JACKING_FDN == 1:											#step 9 jacking down (only jack-up vessel)
                        DUR_JACK_FDN =(PEN_LEG_FDN + AIR_FDN + WD) /JACK_SP_FDN
                        for i in range(0,len(df1.datetime)-start):
                        #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+durationpos-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                            if (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN)+durationjack-1,['s_wht']].max().item() <= HS_JACK): 
                            #print ('good weather window')
                                TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_JACK_FDN
                                break
                            else: 
                            #print('bad weather window')
                                waiting = waiting+1
                                waitingjackfdn = waitingjackfdn+1
                                TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                                i=+1		
                    else:
                        DUR_JACK_FDN = 0 
                    
                    FDN_Installed = FDN_Installed+1
                    FDN_Onboard = FDN_Onboard -1
                    #print('time after installing foundation',FDN_Installed,':',math.ceil(TOTAL_TIME_FDN))
                
                    DUR_REPO_FDN = (INT_D/TRANS_VEL_FDN)                        #step 10 reposition vessel to start next installation 
                    TOTAL_TIME_FDN = TOTAL_TIME_FDN +DUR_REPO_FDN
                    #print('time after repositioning', TOTAL_TIME_FDN)
                    
                else:
                    if stock_FDN > CAP_FDN:										#Check if the stock at the port is sufficent to load vessel 
                        FDN_Onboard = CAP_FDN						#Number of foundations on board is loaded untill vessel capacity is reached 
                    else:
                        FDN_Onboard = stock_FDN									#If stock level is too low load remaining foundations 
                    stock_FDN = stock_FDN - FDN_Onboard	
                    DUR_LOAD_FDN = FDN_Onboard*LT_FDN 							#Step 1 Loading turbines in port 
                    
                    TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_LOAD_FDN
                    #print("Loading complete")
                    #print ("stockfdn",stock_FDN)
                    #print("fdn on deck",FDN_Onboard)
                    
                    #print("total time after loading",math.ceil(TOTAL_TIME_FDN))
          
                   
                    
                
          DUR_TRANS_FDN = (PARK_DIS/TRANS_VEL_FDN)				#Step return to port  		
          for i in range(0,len(df1.datetime)-start):                          #checks wether we are able to set sail with the given limits. 
          
              #print(df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+math.ceil(TOTAL_TIME_FDN+duration-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
              if (df1.loc[start+math.ceil(TOTAL_TIME_FDN):start+1+math.ceil(TOTAL_TIME_FDN)+durationtrans-1,['s_wht']].max().item() <= HS_TRANS_FDN): 
                  #print ('good weather window')
                  TOTAL_TIME_FDN= TOTAL_TIME_FDN + DUR_TRANS_FDN
                  break
              else: 
                 # print('bad weather window')
                  waiting = waiting+1
                  TOTAL_TIME_FDN = TOTAL_TIME_FDN+1
                  i=+1 
                  waitingtransfdn = waitingtransfdn +1
         # print('return to port')          
         # print('TOTAL time installation foundations', math.ceil(TOTAL_TIME_FDN), 'hours')      
          T_FDN.loc[str(year),'Option '+str(option)] = math.ceil(TOTAL_TIME_FDN)
          WAIT_FDN.loc[str(year),'Option '+str(option)] = math.ceil(waiting)
          Waitingtransfdn.loc[str(year),'Option '+str(option)] = math.ceil(waitingtransfdn)
          Waitingposfdn.loc[str(year),'Option '+str(option)] = math.ceil(waitingposfdn)
          Waitingjackfdn.loc[str(year),'Option '+str(option)] = math.ceil(waitingjackfdn)
          Waitingmono.loc[str(year),'Option '+str(option)] = math.ceil(waitingmono)
          Waitingtp.loc[str(year),'Option '+str(option)] = math.ceil(waitingtp)
          #print('Completed foundation installation')
                
        #%% Operation steps Turbine      
        #----- Operation steps  Turbine -----
        WTB_Installed = 0 
        WTB_Onboard = 0 
        waitingwtb = 0 
        start = start +336  #turbines start operation two weeks after foundation installation 
        
        waitingtranswtb = 0 
        waitingposwtb = 0
        waitingjackwtb = 0 
        waitingtow = 0 
        waitingnac = 0 
        waitingblade = 0 
        
        
        TOTAL_TIME_WTB = 0
        stock_WTB = N_WTB 
        
        if SUPPLY_WTB == 0:
         while WTB_Installed<N_WTB:										#Loop to run until all the desired number of turbines are installed 
            if stock_WTB > CAP_WTB:										#Check if the stock at the port is sufficent to load vessel 
                WTB_Onboard = CAP_WTB					                #Number of turbines on board is loaded untill vessel capacity is reached 
            else:
                WTB_Onboard = stock_WTB									#If stock level is too low load remaining turbines
            stock_WTB = stock_WTB - WTB_Onboard							#New stocklevel after loading 
            
           
            DUR_LOAD_WTB = WTB_Onboard*LT_WTB 							#Step 1 Loading turbines in port 
            
            TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_LOAD_WTB
           
            DUR_TRANS_WTB = (PARK_DIS/TRANS_VEL_WTB)			#Step 2 transportation towards site 		
           
            
            durationtranswtb = math.ceil(DUR_TRANS_WTB)                         
            for i in range(0,len(df1.datetime)-start):                          #checks wether we are able to set sail with the given limits. 
            
               
                if (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationtranswtb-1),['s_wht']].max().item() <= HS_TRANS_WTB): 
                 
                    TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_TRANS_WTB
                    break
                else: 
                    
                    waitingwtb = waitingwtb+1
                    TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                    i=+1
                    waitingtranswtb = waitingtranswtb+1
            
            
            while WTB_Onboard>0:                                         #Loop to run untill all turbines on board are installed 
        																
                DUR_POS_WTB = POS_WTB									#Step 3 positioning of vessel 
                #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_POS_WTB
                #print('time after positioning', TOTAL_TIME_WTB)
                durationposwtb = math.ceil(DUR_POS_WTB)                 #calculate the duration of step 2 and 3 to find the length of the weather window we need to find 
                
                for i in range(0,len(df1.datetime)-start):
                    
                    if (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationposwtb-1,['s_wht']].max().item() <= HS_POS_WTB): 
                        #print ('good weather window')
                        TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_POS_WTB 
                        break
                    else: 
                        #print('bad weather window')
                        waitingwtb = waitingwtb+1
                        TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                        i=+1
                        waitingposwtb = waitingposwtb +1
                        
                        
                if JACKING_WTB == 1:											#Step 9 jacking down (only jack-up vessel)
                    DUR_JACK_WTB =(PEN_LEG_WTB + AIR_WTB + WD) /JACK_SP_WTB
                    
                    durationjackwtb = math.ceil(DUR_JACK_WTB)                 #calculate the duration of step 2 and 3 to find the length of the weather window we need to find 
                
                    for i in range(0,len(df1.datetime)-start):
                    
                        if (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationjackwtb-1,['s_wht']].max().item() <= HS_JACK): 
                        #print ('good weather window')
                            TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_JACK_WTB 
                            break
                        else: 
                        #print('bad weather window')
                            waitingwtb = waitingwtb+1
                            TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                            i=+1
                            waitingjackwtb = waitingjackwtb +1 
                        #print('time after jacking down', TOTAL_TIME_WTB) 		
                else:
                    DUR_JACK_WTB = 0
                
                
                
                DUR_HOIST_TOW = N_TOW*LFT_TOW 							        #Step 5 lifting and positioning towerpieces  
                #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_HOIST_TOW
                #print('time after lift tower piece',WTB_Installed+1,':',TOTAL_TIME_WTB)
                
                durationtow= math.ceil(DUR_HOIST_TOW)
                
                for i in range(0,len(df1.datetime)-start):
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationtow-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                    if ((df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationtow-1,['s_wht']].max().item() <= HS_TOW) and (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationtow-1,['wind_speed']].max().item() <= WS_TOW) and (df1.loc[math.ceil(start+TOTAL_TIME_WTB):math.ceil(start+TOTAL_TIME_WTB+durationtow-1),['mean_period']].max().item() <= TP_TOW))  : 
                        #print ('good weather window')
                        TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_HOIST_TOW
                        break
                    else: 
                        #print('bad weather window')
                        waitingwtb = waitingwtb+1
                        TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                        i=+1
                        waitingtow = waitingtow +1 
                
                DUR_HOIST_NAC = LFT_NAC						                    #Step 6 lifting and installing nacelle  
                #TOTAL_TIME_WTB = TOTAL_TIME_WTB +DUR_HOIST_NAC
                #print('time after lift nacelle',WTB_Installed+1,':',TOTAL_TIME_WTB)
                
                durationnac= math.ceil(DUR_HOIST_NAC)
                
                for i in range(0,len(df1.datetime)-start):
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationnac-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                    if ((df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationnac-1,['s_wht']].max().item() <= HS_NAC) and (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationnac-1,['wind_speed']].max().item() <= WS_NAC) and (df1.loc[math.ceil(start+TOTAL_TIME_WTB):math.ceil(start+TOTAL_TIME_WTB+durationnac-1),['mean_period']].max().item() <= TP_NAC))  : 
                        #print ('good weather window')
                        TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_HOIST_NAC
                        break
                    else: 
                        #print('bad weather window')
                        waitingwtb = waitingwtb+1
                        TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                        i=+1
                        waitingnac = waitingnac +1 
                
                
                DUR_INST_BLADE= N_BLADE* (LFT_BLADE + INST_BLADE)		        #Step 7 and 8 Lifting and installing blades 
                #TOTAL_TIME_WTB =TOTAL_TIME_WTB+DUR_INST_BLADE
                #print('time after installing blades',WTB_Installed+1, ':', TOTAL_TIME_WTB)
                
                durationblade= math.ceil(DUR_INST_BLADE)
                
                for i in range(0,len(df1.datetime)-start):
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationblade-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                    if ((df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationblade-1,['s_wht']].max().item() <= HS_BLADE) and (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationblade-1,['wind_speed']].max().item() <= WS_BLADE) and (df1.loc[math.ceil(start+TOTAL_TIME_WTB):math.ceil(start+TOTAL_TIME_WTB+durationblade-1),['mean_period']].max().item() <= TP_BLADE))  : 
                        #print ('good weather window')
                        TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_INST_BLADE
                        break
                    else: 
                        #print('bad weather window')
                        waitingwtb = waitingwtb+1
                        TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                        i=+1
                        waitingblade = waitingblade+1
                        
                if JACKING_WTB == 1:											#Step 9 jacking down (only jack-up vessel)
                    DUR_JACK_WTB =(PEN_LEG_WTB + AIR_WTB + WD) /JACK_SP_WTB
                    #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_JACK_WTB
                    durationjackwtb = math.ceil(DUR_JACK_WTB)                 #calculate the duration of step 2 and 3 to find the length of the weather window we need to find 
                
                    for i in range(0,len(df1.datetime)-start):
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationjackwtb-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                        if (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationjackwtb-1,['s_wht']].max().item() <= HS_JACK): 
                        #print ('good weather window')
                            TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_JACK_WTB 
                            break
                        else: 
                        #print('bad weather window')
                            waitingwtb = waitingwtb+1
                            TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                            i=+1
                            waitingjackwtb = waitingjackwtb+1
                        #print('time after jacking down', TOTAL_TIME_WTB) 		
                else:
                    DUR_JACK_WTB = 0 
                
                
                
                WTB_Installed = WTB_Installed+1
                WTB_Onboard = WTB_Onboard -1
                #print('time after installing turbine',WTB_Installed,':',TOTAL_TIME_WTB)
            
                DUR_REPO_WTB = INT_D/TRANS_VEL_WTB                                #Step 10 reposition the vessel to start next installation 
                TOTAL_TIME_WTB = TOTAL_TIME_WTB +DUR_REPO_WTB
                #print('time after repositioning', TOTAL_TIME_WTB)
            else:
                DUR_TRANS_WTB = PARK_DIS/TRANS_VEL_WTB
				
                durationtranswtb = math.ceil(DUR_TRANS_WTB)                         
                for i in range(0,len(df1.datetime)-start):                          #checks wether we are able to set sail with the given limits. 
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationtranswtb-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                    if (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationtranswtb-1),['s_wht']].max().item() <= HS_TRANS_WTB): 
                        #print ('good weather window')
                        TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_TRANS_WTB
                        break
                    else: 
                        # print('bad weather window')
                        waitingwtb = waitingwtb+1
                        TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                        i=+1
                        waitingtranswtb = waitingtranswtb+1             #Step return to port  		
                #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_TRANS_WTB       
                #print('return to port')
        
        
        elif SUPPLY_WTB == 1:
          TOTAL_TIME_WTB = 0 
          #print('left port')
          DUR_TRANS_WTB = (PARK_DIS/TRANS_VEL_WTB)			#Step 2 transportation towards site 		
          #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_TRANS_WTB
          #print("total time after transit",TOTAL_TIME_WTB)
          durationtranswtb = math.ceil(DUR_TRANS_WTB)                         
          for i in range(0,len(df1.datetime)-start):                          #checks wether we are able to set sail with the given limits. 
              #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationtranswtb-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
              if (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationtranswtb-1),['s_wht']].max().item() <= HS_TRANS_WTB): 
                  #print ('good weather window')
                  TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_TRANS_WTB
                  break
              else: 
                  # print('bad weather window')
                  waitingwtb = waitingwtb+1
                  TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                  i=+1
                  waitingtranswtb = waitingtranswtb+1 
          
          
          
          
          while WTB_Installed<N_WTB:
              
            
                
            while WTB_Onboard>0:                                         #Loop to run untill all turbines on board are installed 
        																
                DUR_POS_WTB = POS_WTB									#Step 3 positioning of vessel 
                #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_POS_WTB
                #print('time after positioning', TOTAL_TIME_WTB)
                
                if JACKING_WTB == 1:											#Step 4 jacking up (only jack-up vessel)
                    DUR_JACK_WTB =(PEN_LEG_WTB + AIR_WTB + WD) /JACK_SP_WTB
                    #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_JACK_WTB
                    #print('time after jacking up', TOTAL_TIME_WTB)			
                else:
                    DUR_JACK_WTB = 0
                
                durationposwtb = math.ceil(DUR_POS_WTB+DUR_JACK_WTB)                 #calculate the duration of step 2 and 3 to find the length of the weather window we need to find 
                
                for i in range(0,len(df1.datetime)-start):
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationposwtb-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                    if (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationposwtb-1,['s_wht']].max().item() <= HS_POS_WTB): 
                        #print ('good weather window')
                        TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_JACK_WTB + DUR_POS_WTB 
                        break
                    else: 
                        #print('bad weather window')
                        waitingwtb = waitingwtb+1
                        TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                        i=+1
                        waitingposwtb = waitingposwtb +1
                        
                DUR_HOIST_TOW = N_TOW*LFT_TOW 							        #Step 5 lifting and positioning towerpieces  
                #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_HOIST_TOW
                #print('time after lift tower piece',WTB_Installed+1,':',TOTAL_TIME_WTB)
                
                durationtow= math.ceil(DUR_HOIST_TOW)
                
                for i in range(0,len(df1.datetime)-start):
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationtow-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                    if ((df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationtow-1,['s_wht']].max().item() <= HS_TOW) and (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationtow-1,['wind_speed']].max().item() <= WS_TOW) and (df1.loc[math.ceil(start+TOTAL_TIME_WTB):math.ceil(start+TOTAL_TIME_WTB+durationtow-1),['mean_period']].max().item() <= TP_TOW))  : 
                        #print ('good weather window')
                        TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_HOIST_TOW
                        break
                    else: 
                        #print('bad weather window')
                        waitingwtb = waitingwtb+1
                        TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                        i=+1
                        waitingtow = waitingtow +1
                
                
                
                DUR_HOIST_NAC = LFT_NAC						                    #Step 6 lifting and installing nacelle  
                #TOTAL_TIME_WTB = TOTAL_TIME_WTB +DUR_HOIST_NAC
                #print('time after lift nacelle',WTB_Installed+1,':',TOTAL_TIME_WTB)
                
                durationnac= math.ceil(DUR_HOIST_NAC)
                
                for i in range(0,len(df1.datetime)-start):
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationnac-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                    if ((df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationnac-1,['s_wht']].max().item() <= HS_NAC) and (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationnac-1,['wind_speed']].max().item() <= WS_NAC) and (df1.loc[math.ceil(start+TOTAL_TIME_WTB):math.ceil(start+TOTAL_TIME_WTB+durationnac-1),['mean_period']].max().item() <= TP_NAC))  : 
                        #print ('good weather window')
                        TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_HOIST_NAC
                        break
                    else: 
                        #print('bad weather window')
                        waitingwtb = waitingwtb+1
                        TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                        i=+1
                        waitingnac = waitingnac+1
                
                
                DUR_INST_BLADE= N_BLADE*(LFT_BLADE + INST_BLADE)		        #Step 7 and 8 Lifting and installing blades 
                #TOTAL_TIME_WTB =TOTAL_TIME_WTB+DUR_INST_BLADE
                #print('time after installing blades',WTB_Installed+1, ':', TOTAL_TIME_WTB)
               
                durationblade= math.ceil(DUR_INST_BLADE)
                
                for i in range(0,len(df1.datetime)-start):
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationblade-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                    if ((df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationblade-1,['s_wht']].max().item() <= HS_BLADE) and (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationblade-1,['wind_speed']].max().item() <= WS_BLADE) and (df1.loc[math.ceil(start+TOTAL_TIME_WTB):math.ceil(start+TOTAL_TIME_WTB+durationblade-1),['mean_period']].max().item() <= TP_BLADE))  : 
                        #print ('good weather window')
                        TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_INST_BLADE
                        break
                    else: 
                        #print('bad weather window')
                        waitingwtb = waitingwtb+1
                        TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                        i=+1
                        waitingblade = waitingblade+1
                
                if JACKING_WTB == 1:											#Step 9 jacking down (only jack-up vessel)
                    DUR_JACK_WTB =(PEN_LEG_WTB + AIR_WTB + WD) /JACK_SP_WTB
                    #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_JACK_WTB
                    durationjackwtb = math.ceil(DUR_JACK_WTB)                 #calculate the duration of step 2 and 3 to find the length of the weather window we need to find 
                
                    for i in range(0,len(df1.datetime)-start):
                    #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationjackwtb-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
                        if (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB)+durationjackwtb-1,['s_wht']].max().item() <= HS_POS_WTB): 
                        #print ('good weather window')
                            TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_JACK_WTB 
                            break
                        else: 
                        #print('bad weather window')
                            waitingwtb = waitingwtb+1
                            TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                            i=+1
                            waitingjackwtb = waitingjackwtb+1 	
                else:
                    DUR_JACK_WTB = 0 
                
                
                WTB_Installed = WTB_Installed+1
                WTB_Onboard = WTB_Onboard -1
                #print('time after installing turbine',WTB_Installed,':',math.ceil(TOTAL_TIME_WTB))
            
                DUR_REPO_WTB = INT_D/TRANS_VEL_WTB                                #Step 10 reposition the vessel to start next installation 
                TOTAL_TIME_WTB = TOTAL_TIME_WTB +DUR_REPO_WTB
                #print('time after repositioning', TOTAL_TIME_WTB)
            else:
                										
                if stock_WTB > CAP_WTB:										#Check if the stock at the port is sufficent to load vessel 
                 WTB_Onboard = CAP_WTB						                #Number of turbines on board is loaded untill vessel capacity is reached 
                else:
                 WTB_Onboard = stock_WTB									#If stock level is too low load remaining turbines
                stock_WTB = stock_WTB - WTB_Onboard							#New stocklevel after loading 
            
                DUR_LOAD_WTB = WTB_Onboard*LT_WTB 							#Step 1 Loading turbines in port 
            
                TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_LOAD_WTB
                #print("Loading complete")
                #print ("stock turbines",stock_WTB)
                #print("turbine on deck after loading",WTB_Onboard)
                #print("total time after loading",math.ceil(TOTAL_TIME_WTB))
          DUR_TRANS_WTB = PARK_DIS/TRANS_VEL_WTB				  #Step return to port  		
          #TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_TRANS_WTB
          durationtranswtb = math.ceil(DUR_TRANS_WTB)                         
          for i in range(0,len(df1.datetime)-start):                          #checks wether we are able to set sail with the given limits. 
              #print(df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationtranswtb-1),['datetime','s_wht', 'wind_speed', 'mean_period']])
              if (df1.loc[start+math.ceil(TOTAL_TIME_WTB):start+math.ceil(TOTAL_TIME_WTB+durationtranswtb-1),['s_wht']].max().item() <= HS_TRANS_WTB): 
                  #print ('good weather window')
                  TOTAL_TIME_WTB= TOTAL_TIME_WTB + DUR_TRANS_WTB
                  break
              else: 
                  # print('bad weather window')
                  waitingwtb = waitingwtb+1
                  TOTAL_TIME_WTB = TOTAL_TIME_WTB+1
                  i=+1
                  waitingtranswtb = waitingtranswtb+1 
                  
                  
        T_WTB.loc[str(year),'Option '+str(option)] = math.ceil(TOTAL_TIME_WTB)
        WAIT_WTB.loc[str(year),'Option '+str(option)] = math.ceil(waitingwtb)
        Waitingtranswtb.loc[str(year),'Option '+str(option)] = math.ceil(waitingtranswtb)
        Waitingposwtb.loc[str(year),'Option '+str(option)] = math.ceil(waitingposwtb)
        Waitingjackwtb.loc[str(year),'Option '+str(option)] = math.ceil(waitingjackwtb)
        Waitingtow.loc[str(year),'Option '+str(option)] = math.ceil(waitingtow)
        Waitingnac.loc[str(year),'Option '+str(option)] = math.ceil(waitingnac)
        Waitingblade.loc[str(year),'Option '+str(option)] = math.ceil(waitingblade)
          #print('return to port')
        #print('Turbine installation completed')    
        #%% Results 
        #------ results-------
        print('   ')
        print('Total time installation foundations', math.ceil(TOTAL_TIME_FDN), 'hours')      
        print('Total time installation turbines', math.ceil(TOTAL_TIME_WTB), 'hours')           
        print('Total waiting time foundations', math.ceil(waiting), 'hours')     
        print('Total waiting time Turbines', math.ceil(waitingwtb), 'hours')
                       
               
        TOTAL_COST_FDN = (TOTAL_TIME_FDN) *(CHT_C_FDN/24)    #Cost of installation of Foundations. Rental cost/hour 
        COST_FDN.loc[str(year),'Option '+str(option)] = math.ceil(TOTAL_COST_FDN)
        
        
        TOTAL_COST_WTB = (TOTAL_TIME_WTB) *(CHT_C_WTB/24)   #Cost of installation of Turbines, Rental cost/hour 
        COST_WTB.loc[str(year),'Option '+str(option)] = math.ceil(TOTAL_COST_WTB)
        
        TOTAL_TIME = TOTAL_TIME_FDN + TOTAL_TIME_WTB
        T_TOT.loc[str(year),'Option '+str(option)] = math.ceil(TOTAL_TIME)
        
        TOTAL_COST_INST = TOTAL_COST_FDN + TOTAL_COST_WTB 
        COST_TOT.loc[str(year),'Option '+str(option)] = math.ceil(TOTAL_COST_INST)
        
        TOTAL_WAIT = waiting + waitingwtb 
        WAIT_TOT.loc[str(year),'Option '+str(option)] = math.ceil(TOTAL_WAIT)
        
        print('   ')
        print('Total cost foundation installation',math.ceil(TOTAL_COST_FDN), 'kDollars')
        print('Total cost turbine installation', math.ceil(TOTAL_COST_WTB), 'kDollars')  
        print('Total Time of installation farm',math.ceil(TOTAL_TIME), 'hours')    
        print('Total Cost of installation farm', math.ceil(TOTAL_COST_INST), 'kDollars')     
        print('Completed farm installation')  
    
  

     