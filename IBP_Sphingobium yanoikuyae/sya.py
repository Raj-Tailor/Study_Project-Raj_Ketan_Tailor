from Bio.KEGG.REST import kegg_get
from Bio.KEGG.KGML.KGML_parser import read
import requests

org = 'sya01100'  # KEGG Organism ID
f = open(org + '.kgml', 'w')  # Creates a file .kgml to save all the information retrieved from KEGG
dat = kegg_get(org, 'kgml')  # Retrieve reactions and compounds from KEGG
datr = dat.read()  # Convert data into a readable format
f.write(datr)  # Write and save
f.close()

p = read(open('sya01100.kgml', 'r'))  # read the model
Metabolites = []  # Starting a list to save all the metabolites

for reaction in p.reactions:  # Walk the reactions. From here you can also extract the ID for all the metabolites
    print('\nReaction')
    print(reaction.name)
    print('Substrates')
    for substrate in reaction.substrates:
        SubstrateName = substrate.name
        SubstrateID = SubstrateName.replace('cpd:', '')
        print(SubstrateID)
        Metabolites.append(SubstrateID)
    print('Products')
    for product in reaction.products:
        ProductName = product.name
        ProductID = ProductName.replace('cpd:', '')
        print(ProductID)
        Metabolites.append(ProductID)

UniqueMetabolites = set(Metabolites)  # Remove duplicates

# Print the number of unique metabolites
print(f'Number of unique metabolites: {len(UniqueMetabolites)}')

url = 'https://www.genome.jp/entry/-f+m+'
for met in UniqueMetabolites:  # This is the list of metabolites name that you can use to retrieve the mol files
    mapa = url + met  # Write a new URL following KEGG format
    print(mapa)
    mol = requests.get(mapa)
    data = mol.text
    x = open(f'{met}.mol', 'w')
    x.write(data)
    x.close()
    print(f'{met}.mol saved successfully')
