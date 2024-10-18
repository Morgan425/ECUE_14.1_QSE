import numpy as np
import matplotlib.pyplot as plt
import openpyxl as xl


charge_efficiency=0.9
sell_efficiency=0.9

sell_price=150
buy_price=20

storage_capacity=10  # MWh
charge_flux=1  
sell_flux=1

Energy_xl=xl.load_workbook('Prix_energie.xlsx').active
Energy_Price=[]
for row in Energy_xl.iter_rows(min_row=2,values_only=True):
    Energy_Price.append(float(row[1]))

temps=len(Energy_Price)


Storage=np.zeros((temps))
Benefit=np.zeros((temps))


for t in range(1,temps):
    price=Energy_Price[t]

    if price<buy_price:
        bought_energy=min(charge_flux*charge_efficiency,storage_capacity-Storage[t-1])
        Storage[t]=Storage[t-1]+bought_energy
        Benefit[t]=Benefit[t-1]-(charge_flux*price)

    elif price>sell_price:
        selled_energy=min(sell_flux,Storage[t-1])
        Storage[t]=Storage[t-1]-selled_energy
        Benefit[t]=Benefit[t-1]+(sell_flux*selled_energy*price)
    
    else:
        Benefit[t]=Benefit[t-1]
        Storage[t]=Storage[t-1]

potential_benefit=round(Benefit[-1]+Storage[-1]*sell_efficiency*Energy_Price[-1],2)
print(f'Potential benefit (if all remaining energy sold): {potential_benefit}€')

fig,(ax1,ax2,ax3)=plt.subplots(3,1,sharex=True)
ax2.plot(Storage)
ax2.set_title('Storage')

ax1.plot(Benefit,'g')
ax1.set_title('Benefit')
ax1.axhline(0,color='black')
ax1.scatter(len(Energy_Price),potential_benefit,color='g')

ax3.plot(Energy_Price,'r')
ax3.set_title('Energy_Price')
plt.show()







