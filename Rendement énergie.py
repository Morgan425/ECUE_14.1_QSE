import numpy as np
import matplotlib.pyplot as plt
import openpyxl as xl


#Parameters

sell_price=120
buy_price=24

charge_efficiency=0.9
sell_efficiency=0.9

initial_storage_percentage=0
storage_capacity=10  # MWh

sell_characteristic=1.2      # Quantité par laquelle la charge est divisée en 1h
constant_buy_flux=1

CC_threshold=0.8

#####


Energy_xl=xl.load_workbook('Prix_energie.xlsx').active
Energy_Price=[]
for row in Energy_xl.iter_rows(min_row=2,values_only=True):
    Energy_Price.append(float(row[1]))

temps=len(Energy_Price)


Storage=np.empty((temps))
Benefit=np.empty((temps))
Benefit[0]=0
Storage[0]=storage_capacity*initial_storage_percentage


for t in range(1,temps):
    price=Energy_Price[t]

    if price<buy_price:
        if Storage[t-1]<CC_threshold*storage_capacity:
            bought_energy=min(constant_buy_flux*charge_efficiency,storage_capacity-Storage[t-1])
        else:
            remaining_storage_percentage=(storage_capacity-Storage[t-1])/(storage_capacity*(1-CC_threshold))
            bought_energy=constant_buy_flux*charge_efficiency*remaining_storage_percentage
        Storage[t]=Storage[t-1]+bought_energy
        Benefit[t]=Benefit[t-1]-((bought_energy*price)/charge_efficiency)

    elif price>sell_price:
        Storage[t]=round(Storage[t-1]/sell_characteristic,3)
        selled_energy=Storage[t-1]-Storage[t]
        Benefit[t]=Benefit[t-1]+selled_energy*price*sell_efficiency
    
    else:
        Benefit[t]=Benefit[t-1]
        Storage[t]=Storage[t-1]

potential_benefit=round(Benefit[-1]+Storage[-1]*sell_efficiency*Energy_Price[-1],2)
print(f'Potential benefit (if all remaining energy sold): {potential_benefit}€')


fig,(ax1,ax2,ax3)=plt.subplots(3,1,sharex=True)

ax1.plot(Benefit,'g')
ax1.set_title('Benefit')
ax1.axhline(0,color='black')
ax1.scatter(len(Energy_Price),potential_benefit,color='darkgreen')
ax1.text(len(Energy_Price)*0.9,potential_benefit*1.1,'Potential Benefit',color='darkgreen',fontsize=7,fontweight='bold')
ax1.grid('on')


ax2.plot(Storage)
ax2.set_title('Storage')
ax2.grid('on')

ax3.plot(Energy_Price,'orange')
ax3.set_title('Energy_Price')
ax3.grid('on')
ax3.axhline(sell_price,color='r')
ax3.axhline(buy_price,color='g')

plt.show()



