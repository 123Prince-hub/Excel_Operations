import pandas as pd
a = pd.DataFrame()
d = {'id': [1, 2, 10, 12],
	'val1': ['a', 'b', 'c', 'd']}

a = pd.DataFrame(d)
print(a)

d = {'id': [1, 2, 9, 8],
     'val1': ['p', 'q', 'r', 's']}
b = pd.DataFrame(d)
print(b)
