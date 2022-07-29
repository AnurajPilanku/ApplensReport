import os
import sys
path=sys.argv[1]
files=os.listdir(path)
if len(files)==0:
    print("failure")
else:
    os.rename(os.path.join(path,files[0]),os.path.join(path,"applens.xlsx"))
    print("success")