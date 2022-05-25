import matplotlib.pyplot as plt
import pandas as pd

loc3 = ("Results calm weather.xlsx") 
Results_Calm = pd.read_excel(loc3, index_col=('Configurations'))


Mean_Operability=Results_Calm.T_TOT/T_TOT.mean()

T_TOT.plot(kind='box', rot=45, grid=True, figsize = (15,10), fontsize=20, legend=True)
plt.bar(range(1,len(Results_Calm)+1),Results_Calm.T_TOT,label='Calm weather Operations')
plt.ylabel('Total duration farm installation [h]', fontsize= 20)
plt.legend(fontsize= 'xx-large', frameon=False, loc = 'upper center')
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Total duration farm installation.pdf',bbox_inches='tight')
#plt.title("Total duration installation offshore wind farm, calm weather comparison", fontsize = 28)


COST_TOT.plot(kind='box', rot=45, grid=True, figsize = (15,10), fontsize=20)
plt.bar(range(1,len(Results_Calm)+1),Results_Calm.COST_TOT,label='Calm weather Operations')
plt.ylabel('Total Cost farm installation [k€]', fontsize=20)
plt.legend(fontsize = 'xx-large', frameon=False)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Total cost farm installation.pdf',bbox_inches='tight')
#plt.title("Total cost installation offshore wind farm, calm weather comparison", fontsize = 28)


T_FDN.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20, ylim=(0,1800))
plt.bar(range(1,len(Results_Calm)+1),Results_Calm.T_FDN,label='Calm weather Operations')
plt.ylabel('Total duration Foundation Installation [h]', fontsize = 20)
plt.legend(fontsize = 'xx-large', frameon=False)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Total duration foundation installation.pdf',bbox_inches='tight')
#plt.title("Duration foundation Installation, Calm weather comparison", fontsize = 28)



T_WTB.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20, ylim=(0,1800))
#plt.xlabel("Configuration Options", fontsize =20)
plt.bar(range(1,len(Results_Calm)+1),Results_Calm.T_WTB,label='Calm weather Operations')
plt.ylabel('Total duration Turbine Installation [h]', fontsize = 20)
plt.legend(fontsize = 'xx-large', frameon=False, loc ='upper right')
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Total duration turbine installation.pdf',bbox_inches='tight')
#plt.title("Duration turbine Installation, Calm weather comparison", fontsize = 28)


COST_FDN.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20, ylim=(0,30000))
plt.bar(range(1,len(Results_Calm)+1),Results_Calm.COST_FDN,label='Calm weather Operations')
plt.ylabel('Total cost Foundation Installation [k€]', fontsize = 20)
plt.legend(fontsize = 'xx-large', frameon=False, loc ='upper right')
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Total cost foundation installation.pdf',bbox_inches='tight')
#plt.title("Duration turbine Installation, Calm weather comparison", fontsize = 28)


COST_WTB.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20, ylim=(0,30000))
plt.bar(range(1,len(Results_Calm)+1),Results_Calm.COST_WTB,label='Calm weather Operations')
plt.ylabel('Total cost Turbine Installation [k€]', fontsize = 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Total cost turbine installation.pdf',bbox_inches='tight')
plt.legend(fontsize = 'xx-large', frameon=False, loc ='upper center')
#plt.title("Duration turbine Installation, Calm weather comparison", fontsize = 28)


plt.figure()
WAIT_TOT.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total duration waiting on weather [h]', fontsize = 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Total waiting farm installation.pdf',bbox_inches='tight')
#plt.title('Total duration waiting on weather offshore windfarm installation', fontsize = 28)


plt.figure()
Mean_Operability.plot(kind = 'bar', rot=45, grid = True, figsize= (12,10), fontsize = 20 )
plt.ylabel('Operability factor', fontsize = 20)
plt.xlabel(" ", fontsize =1)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Operability factor.pdf', bbox_inches = 'tight')


Avg_Increasepdtime = 100*((T_TOT.mean() - Results_Calm.loc[:,'T_TOT'])/Results_Calm.loc[:,'T_TOT'])  #avg increase in total duration per Option 
Avg_Increasepdcost =100*((COST_TOT.mean() - Results_Calm.loc[:,'COST_TOT']))/(Results_Calm.loc[:,'COST_TOT'])
    
plt.figure()
Avg_Increasepdcost.plot(kind= 'bar', rot =45,  grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Average increase in Cost of installation [%]', fontsize = 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Average increase cost.pdf', bbox_inches= 'tight')


plt.figure()
Avg_Increasepdtime.plot(kind= 'bar', rot= 45,  grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Average increase in Duration of installation [%]', fontsize = 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on() 
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.2)
plt.savefig('Average increase duration.pdf', bbox_inches  = 'tight')
#%% avg cost and duration install 

Avg_Install = T_TOT.mean()/720
Avg_InstallCalm = Results_Calm.T_TOT/720

plt.figure()
Avg_Install.plot(kind= 'bar', figsize=(20,10), fontsize = 20 , grid =True)
plt.ylabel('Average duration wind turbine install [days/assemby]', fontsize = 22)
#plt.title('Average duration of installation per completed assembly', fontsize= 28)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
plt.savefig('Average duration op installation per assembly.pdf', bbox_inches= 'tight')


Avg_Cost = COST_TOT.mean()/30
plt.figure()
Avg_Cost.plot(kind= 'bar', figsize=(20,10), fontsize = 20 , grid =True)
plt.ylabel('Average Cost wind turbine install [k€/MW]', fontsize = 22)
#plt.title('Average Cost of installation per completed assembly', fontsize= 28)
#plt.xlabel("Configuration Options", fontsize =20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
plt.savefig('Average cost op installation per assembly.pdf', bbox_inches ='tight')


#%% individ waiting
#individual waiting plots fdn 
Waitingtransfdn.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=19, ylim = (0,175))
plt.ylabel('Total waiting transit foundation installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting transportation foundation installation', fontsize = 28)
plt.savefig('Total waiting on transit fdn.pdf',bbox_inches='tight')

Waitingposfdn.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total waiting positioning foundation installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting positioning foundation installation', fontsize = 28)
plt.savefig('Total waiting on positioning fdn.pdf',bbox_inches='tight')

Waitingjackfdn.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total waiting jacking foundation installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting jacking foundation installation', fontsize = 28)
plt.savefig('Total waiting on jacking fdn.pdf',bbox_inches='tight')

Waitingmono.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total waiting monopile installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting monopile installation', fontsize = 28)
plt.savefig('Total waiting on monopile.pdf',bbox_inches='tight')

Waitingtp.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total waiting transition piece installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting transition piece installation', fontsize = 28)
plt.savefig('Total waiting on transition piece.pdf',bbox_inches='tight')

#individual waiting plots wtb 

Waitingtranswtb.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20, ylim = (0,145))
plt.ylabel('Total waiting transportation turbine installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting transportation turbine installation', fontsize = 28)
plt.savefig('Total waiting on transit wtb.pdf',bbox_inches='tight')

Waitingposwtb.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total waiting positioning turbine installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting positioning turbine installation', fontsize = 28)
plt.savefig('Total waiting on positioning wtb.pdf',bbox_inches='tight')

Waitingjackwtb.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total waiting jacking turbine installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#.title('Total duration waiting jacking turbine installation', fontsize = 28)
plt.savefig('Total waiting on jacking wtb.pdf',bbox_inches='tight')

Waitingtow.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total waiting tower installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting tower installation', fontsize = 28)
plt.savefig('Total waiting on tower.pdf',bbox_inches='tight')

Waitingnac.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total waiting nacelle installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting nacelle installation', fontsize = 28)
plt.savefig('Total waiting on nacelle.pdf',bbox_inches='tight')

Waitingblade.plot(kind='box', rot=45, grid=True, figsize = (12,10), fontsize=20)
plt.ylabel('Total waiting blade installation [h]', fontsize= 20)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
#plt.title('Total duration waiting blade installation', fontsize = 28)
plt.savefig('Total waiting on blade.pdf',bbox_inches='tight')


#%%
Avg_Installfdn = T_FDN.mean()/720


plt.figure()
Avg_Installfdn.plot(kind= 'bar', figsize=(12,10), rot = 45, fontsize = 20 , grid =True)
plt.ylabel('Average duration foundation installation [days/foundation]', fontsize = 22)
#plt.title('Average duration of installation per completed assembly', fontsize= 28)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
plt.savefig('Average duration op installation per foundation.pdf', bbox_inches= 'tight')


Avg_Installwtb = T_WTB.mean()/720


plt.figure()
Avg_Installwtb.plot(kind= 'bar', figsize=(12,10), rot=45, fontsize = 20 , grid =True)
plt.ylabel('Average duration turbine installation [days/turbine]', fontsize = 22)
#plt.title('Average duration of installation per completed assembly', fontsize= 28)
plt.grid(b=True, which='major', color='#666666', linestyle='-')
plt.minorticks_on()
plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
plt.savefig('Average duration op installation per turbine.pdf', bbox_inches= 'tight')


# Avg_Cost = COST_TOT.mean()/30
# plt.figure()
# Avg_Cost.plot(kind= 'bar', figsize=(20,10), fontsize = 20 , grid =True)
# plt.ylabel('Average Cost wind turbine install [k€/MW]', fontsize = 22)
# #plt.title('Average Cost of installation per completed assembly', fontsize= 28)
# #plt.xlabel("Configuration Options", fontsize =20)
# plt.grid(b=True, which='major', color='#666666', linestyle='-')
# plt.minorticks_on()
# plt.grid(b=True, which='minor', color='#999999', linestyle='-', alpha=0.1)
# plt.savefig('Average cost op installation per assembly.pdf', bbox_inches ='tight')