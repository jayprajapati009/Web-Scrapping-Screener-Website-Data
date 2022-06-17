a = ['1,75', '1,h', '1']
b = ['12,5', '1,h', '1']
c = ['12,5', '1,h', '1']
d = ['12,5', '1,h', '1']

h = [a, b, c, d]

for i in h:
    for j in i:
        i[i.index(j)] = j.replace(',', '')

print(h)