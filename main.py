import hello
import time
import interaction_sys
import pandas as pd

hello.hello()
print(interaction_sys.get_file_path())
for elem in time.gmtime():
    print(elem)

data = pd.DataFrame({
    "score" : [99, 88, 77, 66]
})
data.to_excel(interaction_sys.get_file_path())