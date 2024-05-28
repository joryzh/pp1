def data_wash(pd, col):
    sum = pd.loc[1:,col].sum()
    avg = pd.loc[1:,col].mean()
    min = pd.loc[1:,col].min()
    min_idx = pd.loc[1:,col].idxmin()
    max = pd.loc[1:,col].max()
    max_idx = pd.loc[1:,col].idxmax()
    # print(pd.loc[1:,1],sum,avg,max,max_idx, min, min_idx)
    print('name: %s, column: %d, sum: %d, avg: %f, max: %d, idx: %d, min: %d, idx: %d' 
          % (pd.loc[0,col],col,sum,avg,max,max_idx, min, min_idx))
    return sum