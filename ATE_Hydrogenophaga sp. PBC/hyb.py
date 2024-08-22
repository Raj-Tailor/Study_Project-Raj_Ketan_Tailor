from Bio.KEGG.REST import kegg_get
from Bio.KEGG.KGML.KGML_parser import read
import requests

org = 'hyb01100'  # KEGG Organism ID
f = open(org + '.kgml', 'w')  # Creates .kgml file and saves all the information retrieved from KEGG
dat = kegg_get(org, 'kgml')  # Retrieve reactions and compounds from KEGG
datr = dat.read()  # Convert data into a readable format
f.write(datr)  # Write and save
f.close()

p = read(open('hyb01100.kgml', 'r'))  # reads the model
Metabolites = []  # Creates a list to save all the metabolites

for reaction in p.reactions:  # iterating the reactions for extracting the ID's of all the metabolites
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
for met in UniqueMetabolites:  # List of metabolites name used to retrieve the mol files
    mapa = url + met  # Constructing new URL according to KEGG format
    print(mapa)
    mol = requests.get(mapa)
    data = mol.text
    x = open(f'{met}.mol', 'w')
    x.write(data)
    x.close()
    print(f'{met}.mol saved successfully')
