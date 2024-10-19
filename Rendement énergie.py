import numpy as np
import matplotlib.pyplot as plt
import openpyxl as xl


#### Parameters ############################################################################################################

#sell_price=83
#buy_price=21

charge_efficiency=0.95          # efficacité du processus de charge de batterie
discharge_efficiency=0.95       # efficacité du processus de décharge de batterie

initial_storage_percentage=0    # énergie initiale dans la batterie (%)
storage_capacity=10             # Capacité tootale de la batterie (MWh)

sell_characteristic=1.2         # Quantité par laquelle la charge est divisée en 1h
constant_buy_flux=1             # Vitesse de charge de la batterie pendant l'étape de charge continue

CC_threshold=0.8                # Niveau de charge entre les 2 étapes du chargement

mean_number=10                  # Période sur laquelle on fait la moyenne pour la décision d'achat ou de vente

############################################################################################################################

Energy_xl=xl.load_workbook('Prix_energie.xlsx').active        #
Energy_Price=[]                                               # Récupération des données du prix de lélectricité
for row in Energy_xl.iter_rows(min_row=2,values_only=True):   #
    Energy_Price.append(float(row[1]))                        #


Energy_prices_mean=[]                           # Moyennage de l'énergie sur 2*{mean_number}
for i in range(len(Energy_Price)):              #
    Energy_prices_mean.append(np.mean(Energy_Price[max(0,i - mean_number):min(i + mean_number,len(Energy_Price))]))


temps=len(Energy_Price)

Storage=np.empty((temps))                                # Evolution du stockage au cours du temps
Benefit=np.empty((temps))                                # Evolution du bénéfice au cours du temps
Benefit[0]=0                                             # Initialisation du bénéfice
Storage[0]=storage_capacity * initial_storage_percentage   # Initialisation du stockage selon {initial_storage_percentage}


for t in range(1,temps):

    price=Energy_Price[t]

    if price < Energy_prices_mean[t] * charge_efficiency:       # Cas d'achat: l'énergie effectivement achetée est à un prix inférieur à sa moyenne sur une plus grande période

        if Storage[t-1] < CC_threshold * storage_capacity:          # Etape de charge constante
            bought_energy=min(constant_buy_flux * charge_efficiency,storage_capacity-Storage[t-1])

        else:                                                       # Etape de fin de charge
            remaining_storage_percentage=(storage_capacity-Storage[t-1])/(storage_capacity * (1-CC_threshold))
            bought_energy=constant_buy_flux * charge_efficiency * remaining_storage_percentage
        Storage[t]=Storage[t-1] + bought_energy
        Benefit[t]=Benefit[t-1] - ((bought_energy * price)/charge_efficiency)


    elif price > Energy_prices_mean[t] / discharge_efficiency:  # Cas de vente: l'énergie effectivement vendue est à un prix supérieur à sa moyenne sur une plus grande période
        Storage[t]=round(Storage[t-1] / sell_characteristic,3)
        selled_energy=Storage[t-1] - Storage[t]
        Benefit[t]=Benefit[t-1] + selled_energy * price * discharge_efficiency
    
    else:                                                       # Aucune action effectuée
        Benefit[t]=Benefit[t-1]
        Storage[t]=Storage[t-1]

potential_benefit=round(Benefit[-1] + Storage[-1] * discharge_efficiency * Energy_Price[-1],2)  # Calcul du bénéfice potentiel en rajoutant au bénéfice la valeur finale de l'énergie stockée
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
#ax3.axhline(sell_price,color='r')
#ax3.axhline(buy_price,color='g')

ax3.plot(Energy_prices_mean,'purple')


plt.show()



