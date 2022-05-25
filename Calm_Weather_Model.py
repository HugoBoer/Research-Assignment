# -*- coding: utf-8 -*-
# %% File reading
"""
Created on Wed Dec  1 15:15:30 2021

@author: Hugo Boer
"""
import pandas as pd  
import math

loc=("Vessel Data.xls")
loc2= ("Farm Data.xls")


Data_Vessel = pd.read_excel(loc, sheet_name = 'Vessel')	     #Vessel data sheet 
Data_Operation =pd.read_excel(loc,sheet_name ='Operation')   #Operation data sheet 
Data_Farm = pd.read_excel(loc2) 		     #Farm data sheet
Data_Config = pd.read_excel(loc, sheet_name = 'Configurations')

for option in range (1,10):
    
    VT_FDN = Data_Config.iloc[option,1]                #Vessel type used for foundation installation  0 = Crane barge, 1= HLV, 2 = Jack-up 
    VT_WTB = Data_Config.iloc[option,2]                #Vessel type used for turbine installation, 0= Crane barge, 1=HLV, 2= Jack-up 
    
    
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
    PARK_DIS = Data_Farm.iloc[9,2]						#Distance farm to feeding harbour 
    
    
    #--------------Vessel Characteristics----- 
    
    #Foundation vessel 
    
    TRANS_VEL_FDN= Data_Vessel.iloc[2,2+VT_FDN]               #Transition Velocity of selected Foundation installation Vessel 
    CAP_FDN= Data_Vessel.iloc[3,2+VT_FDN]						#Loading capacity Foundation of selected Foundation installation Vessel 
    HAM_SP = Data_Vessel.iloc[5,2+VT_FDN]	                #Hamering Speed of selected Foundation Installation Vessel 			
    AIR_FDN = Data_Vessel.iloc[7,2+VT_FDN]					#Air gap between vessel and sea level, only for jack-up Vessels 
    PEN_LEG_FDN = Data_Vessel.iloc[8,2+VT_FDN]				#Penetration depth of the legs of jack-up vessel 
    JACK_SP_FDN = Data_Vessel.iloc[9,2+VT_FDN]				#Jack-up-speed, only for jack-up vessels
    CHT_C_FDN =  Data_Vessel.iloc[12,2+VT_FDN]				#Day-rate charter cost FDN vessel 
    JACKING_FDN =Data_Vessel.iloc[14,2+VT_FDN]               #Does the vessel require jacking 
    SUPPLY_FDN = Data_Vessel.iloc[15,2+VT_FDN]
    
    
    
    #Turbine installation vessel 
    TRANS_VEL_WTB= Data_Vessel.iloc[2,2+VT_WTB]               #Transition Velocity WTB Vessel 
    CAP_WTB = Data_Vessel.iloc[4,2+VT_WTB]					#Loading capacity WTB vessel 
    AIR_WTB = Data_Vessel.iloc[7,2+VT_WTB]					#Air gap between vessel and sea level, only for jack-up Vessels 
    PEN_LEG_WTB =Data_Vessel.iloc[8,2+VT_WTB]				#Penetration depth of the legs of jack-up vessel 
    JACK_SP_WTB = Data_Vessel.iloc[9,2+VT_WTB]				#Jack-up-speed, only for jack-up vessels
    CHT_C_WTB =  Data_Vessel.iloc[12,2+VT_WTB]				#Day-rate charter cost selected WTB vessel 
    JACKING_WTB = Data_Vessel.iloc[14,2+VT_WTB]               #Does the vessel require jacking 
    SUPPLY_WTB =Data_Vessel.iloc[15,2+VT_WTB]
    
    # #---------Operation Characteristics -----
    
    LT_FDN 		= Data_Operation.iloc[2,2+VT_FDN]			        #Loading time foundations 
    LT_WTB 		= Data_Operation.iloc[3,2+VT_WTB]				    #Loading time Turbine components (includes towers, nacelles and blades)
    POS_FDN 	= Data_Operation.iloc[4,2+VT_FDN]			#Postitioning Time of selected Foundation Installation Vessel 
    POS_WTB 	= Data_Operation.iloc[4,2+VT_WTB]			#Postiitioning Time of selected Wind Turbine installation vessel 
    LFT_FDN 	= Data_Operation.iloc[5,2+VT_FDN]			#Lifting time foundation 
    LFT_TP		= Data_Operation.iloc[6,2+VT_FDN]		#Lifting time transition piece 
    GRT_TP 		= Data_Operation.iloc[7,2+VT_FDN]			#Grouting time transition piece 
    LFT_TOW 	= Data_Operation.iloc[8,2+VT_WTB]			#Lifting time tower piece 
    LFT_NAC 	= Data_Operation.iloc[9,2+VT_WTB]			#Lifting time nacelle 
    LFT_BLADE	= Data_Operation.iloc[10,2+VT_WTB]			#lifting time blade 
    INST_BLADE 	= Data_Operation.iloc[11,2+VT_WTB]		#installation time blade 
    
    #%% Operation steps Foundation 
    TOTAL_TIME_FDN = 0 
    stock_FDN = N_FDN 
    
    
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
            #print ("stockfdn",stock_FDN)
            #print("fdn on deck after port",FDN_Onboard)
            #print("Left port")
            #print("total time after loading",TOTAL_TIME_FDN)
            DUR_TRANS_FDN = (PARK_DIS/TRANS_VEL_FDN)			#Step 2 transportation towards site 		
            TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_TRANS_FDN
            #print("total time after transit",TOTAL_TIME_FDN)
                
            while FDN_Onboard>0:
        																#Loop to run untill all foundation on board are installed 
                DUR_POS_FDN = POS_FDN									#Step 3 positioning of vessel 
                TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_POS_FDN
                #print('time after positioning', TOTAL_TIME_FDN)
                
                if JACKING_FDN == 1:											#step 4 jacking up (only jack-up vessel)
                    DUR_JACK_FDN =(PEN_LEG_FDN + AIR_FDN + WD) /JACK_SP_FDN
                    TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_JACK_FDN
                    #print('time after jacking up', TOTAL_TIME_FDN)			
                else:
                    DUR_JACK_FDN = 0
                
                DUR_HOIST_FDN = LFT_FDN 									#step 5 lifting and positioning monopile on seabed 
                TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_HOIST_FDN
                #print('time after lift MP',FDN_Installed+1,':',TOTAL_TIME_FDN)
                
                DUR_HAM = PEN_FDN/HAM_SP          						#step 6 Hammer monopile down using hydraulic jack 
                TOTAL_TIME_FDN = TOTAL_TIME_FDN +DUR_HAM
                #print('time after hammering MP',FDN_Installed+1,':',TOTAL_TIME_FDN)
                
                DUR_HOIST_TP = LFT_TP 									#Step 7 Lifting Transition piece into place 
                TOTAL_TIME_FDN =TOTAL_TIME_FDN+DUR_HOIST_TP
                #print('time after lifting TP',FDN_Installed+1, ':', TOTAL_TIME_FDN)
                
                DUR_INST_TP = GRT_TP									#Step 8 Installing/Grouting Transition Piece 
                TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_INST_TP
                #print('time after grouting TP',FDN_Installed+1,':',TOTAL_TIME_FDN)
                
                if JACKING_FDN == 1:											#step 9 jacking down (only jack-up vessel)
                    DUR_JACK_FDN =(PEN_LEG_FDN + AIR_FDN + WD) /JACK_SP_FDN
                    TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_JACK_FDN
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
                TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_TRANS_FDN        
               # print('return to port')
        #print('TOTAL time installation foundations', math.ceil(TOTAL_TIME_FDN), 'hours')      
            
        
        #print('Completed foundation installation')
        
    elif SUPPLY_FDN == 1:
        
      #print("Left port")
      DUR_TRANS_FDN = (PARK_DIS/TRANS_VEL_FDN)			#Step 2 transportation towards site 		
      TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_TRANS_FDN
     # print("total time after transit",TOTAL_TIME_FDN)  
        
        
        
      while FDN_Installed<N_FDN:										#Loop to run until all the desired number of turbines are installed 
            						
            
                
            while FDN_Onboard>0:
        																#Loop to run untill all foundation on board are installed 
                DUR_POS_FDN = POS_FDN									#Step 3 positioning of vessel 
                TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_POS_FDN
                #print('time after positioning', TOTAL_TIME_FDN)
                
                if JACKING_FDN == 1:											#step 4 jacking up (only jack-up vessel)
                    DUR_JACK_FDN =(PEN_LEG_FDN + AIR_FDN + WD) /JACK_SP_FDN
                    TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_JACK_FDN
                    #print('time after jacking up', TOTAL_TIME_FDN)			
                else:
                    DUR_JACK_FDN = 0
                
                DUR_HOIST_FDN = LFT_FDN 									#step 5 lifting and positioning monopile on seabed 
                TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_HOIST_FDN
                #print('time after lift MP',FDN_Installed+1,':',TOTAL_TIME_FDN)
                
                DUR_HAM = PEN_FDN/HAM_SP          						#step 6 Hammer monopile down using hydraulic jack 
                TOTAL_TIME_FDN = TOTAL_TIME_FDN +DUR_HAM
                #print('time after hammering MP',FDN_Installed+1,':',TOTAL_TIME_FDN)
                
                DUR_HOIST_TP = LFT_TP 									#Step 7 Lifting Transition piece into place 
                TOTAL_TIME_FDN =TOTAL_TIME_FDN+DUR_HOIST_TP
                #print('time after lifting TP',FDN_Installed+1, ':', TOTAL_TIME_FDN)
                
                DUR_INST_TP = GRT_TP									#Step 8 Installing/Grouting Transition Piece 
                TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_INST_TP
                #print('time after grouting TP',FDN_Installed+1,':',TOTAL_TIME_FDN)
                
                if JACKING_FDN == 1:											#step 9 jacking down (only jack-up vessel)
                    DUR_JACK_FDN =(PEN_LEG_FDN + AIR_FDN + WD) /JACK_SP_FDN
                    TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_JACK_FDN
                    #print('time after jacking down', TOTAL_TIME_FDN) 		
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
      
                DUR_TRANS_FDN = (PARK_DIS/TRANS_VEL_FDN)
      DUR_TRANS_FDN = (PARK_DIS/TRANS_VEL_FDN)				#Step return to port  		
      TOTAL_TIME_FDN = TOTAL_TIME_FDN + DUR_TRANS_FDN        
      #print('return to port')          
     # print('TOTAL time installation foundations', math.ceil(TOTAL_TIME_FDN), 'hours')      
            
      #print('Completed foundation installation')
            
    #%% Operation steps Turbine      
    #----- Operation steps  Turbine -----
    WTB_Installed = 0 
    WTB_Onboard = 0 
    
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
        #print("Loading complete")
        #print ("stock turbines",stock_WTB)
        #print("turbine on deck after port",WTB_Onboard)
        #print("Left port")
        #print("total time after loading",TOTAL_TIME_WTB)
        DUR_TRANS_WTB = (PARK_DIS/TRANS_VEL_WTB)			#Step 2 transportation towards site 		
        TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_TRANS_WTB
        #print("total time after transit",TOTAL_TIME_WTB)
            
        while WTB_Onboard>0:                                         #Loop to run untill all turbines on board are installed 
    																
            DUR_POS_WTB = POS_WTB									#Step 3 positioning of vessel 
            TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_POS_WTB
            #print('time after positioning', TOTAL_TIME_WTB)
            
            if JACKING_WTB == 1:											#Step 4 jacking up (only jack-up vessel)
                DUR_JACK_WTB =(PEN_LEG_WTB + AIR_WTB + WD) /JACK_SP_WTB
                TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_JACK_WTB
                #print('time after jacking up', TOTAL_TIME_WTB)			
            else:
                DUR_JACK_WTB = 0
            
            
            DUR_HOIST_TOW = N_TOW*LFT_TOW 							        #Step 5 lifting and positioning towerpieces  
            TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_HOIST_TOW
            #print('time after lift tower piece',WTB_Installed+1,':',TOTAL_TIME_WTB)
            
            DUR_HOIST_NAC = LFT_NAC						                    #Step 6 lifting and installing nacelle  
            TOTAL_TIME_WTB = TOTAL_TIME_WTB +DUR_HOIST_NAC
            #print('time after lift nacelle',WTB_Installed+1,':',TOTAL_TIME_WTB)
            
            
            
            DUR_INST_BLADE= N_BLADE* (LFT_BLADE + INST_BLADE)		        #Step 7 and 8 Lifting and installing blades 
            TOTAL_TIME_WTB =TOTAL_TIME_WTB+DUR_INST_BLADE
            #print('time after installing blades',WTB_Installed+1, ':', TOTAL_TIME_WTB)
           
            
            if JACKING_WTB == 1:											#Step 9 jacking down (only jack-up vessel)
                DUR_JACK_WTB =(PEN_LEG_WTB + AIR_WTB + WD) /JACK_SP_WTB
                TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_JACK_WTB
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
            DUR_TRANS_WTB = PARK_DIS/TRANS_VEL_WTB				  #Step return to port  		
            TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_TRANS_WTB       
            #print('return to port')
    
    
    elif SUPPLY_WTB == 1:
        
      #print('left port')
      DUR_TRANS_WTB = (PARK_DIS/TRANS_VEL_WTB)			#Step 2 transportation towards site 		
      TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_TRANS_WTB
      #print("total time after transit",TOTAL_TIME_WTB)
      
      while WTB_Installed<N_WTB:
          
        
            
        while WTB_Onboard>0:                                         #Loop to run untill all turbines on board are installed 
    																
            DUR_POS_WTB = POS_WTB									#Step 3 positioning of vessel 
            TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_POS_WTB
            #print('time after positioning', TOTAL_TIME_WTB)
            
            if JACKING_WTB == 1:											#Step 4 jacking up (only jack-up vessel)
                DUR_JACK_WTB =(PEN_LEG_WTB + AIR_WTB + WD) /JACK_SP_WTB
                TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_JACK_WTB
                #print('time after jacking up', TOTAL_TIME_WTB)			
            else:
                DUR_JACK_WTB = 0
            
            
            DUR_HOIST_TOW = N_TOW*LFT_TOW 							        #Step 5 lifting and positioning towerpieces  
            TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_HOIST_TOW
            #print('time after lift tower piece',WTB_Installed+1,':',TOTAL_TIME_WTB)
            
            DUR_HOIST_NAC = LFT_NAC						                    #Step 6 lifting and installing nacelle  
            TOTAL_TIME_WTB = TOTAL_TIME_WTB +DUR_HOIST_NAC
            #print('time after lift nacelle',WTB_Installed+1,':',TOTAL_TIME_WTB)
            
            
            
            DUR_INST_BLADE= N_BLADE*(LFT_BLADE + INST_BLADE)		        #Step 7 and 8 Lifting and installing blades 
            TOTAL_TIME_WTB =TOTAL_TIME_WTB+DUR_INST_BLADE
            #print('time after installing blades',WTB_Installed+1, ':', TOTAL_TIME_WTB)
           
            
            if JACKING_WTB == 1:											#Step 9 jacking down (only jack-up vessel)
                DUR_JACK_WTB =(PEN_LEG_WTB + AIR_WTB + WD) /JACK_SP_WTB
                TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_JACK_WTB
                #print('time after jacking down', TOTAL_TIME_WTB) 		
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
      TOTAL_TIME_WTB = TOTAL_TIME_WTB + DUR_TRANS_WTB       
      #print('return to port')
    #print('Turbine installation completed')    
    
    #%% Results 
    #------ results-------
    print('   ')
    print('Total time installation foundations', math.ceil(TOTAL_TIME_FDN), 'hours')      
    print('Total time installation turbines', math.ceil(TOTAL_TIME_WTB), 'hours')           
            
    
                   
           
    TOTAL_COST_FDN = (TOTAL_TIME_FDN) *(CHT_C_FDN/24)    #Cost of installation of Foundations. Rental cost/hour 
    
    TOTAL_COST_WTB = (TOTAL_TIME_WTB) *(CHT_C_WTB/24)   #Cost of installation of Turbines, Rental cost/hour 
    
    TOTAL_TIME = TOTAL_TIME_FDN + TOTAL_TIME_WTB
    TOTAL_COST_INST = TOTAL_COST_FDN + TOTAL_COST_WTB
    
    print('   ')
    print('Total cost foundation installation',math.ceil(TOTAL_COST_FDN), 'kDollars')
    print('Total cost turbine installation', math.ceil(TOTAL_COST_WTB), 'kDollars')  
    print('Total Time of installation farm',math.ceil(TOTAL_TIME), 'hours')    
    print('Total Cost of installation farm', math.ceil(TOTAL_COST_INST), 'kDollars')     
        