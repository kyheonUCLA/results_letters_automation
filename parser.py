import pandas as pd
import numpy as np

class DataParser():
    def __init__(self, filepath) -> None:
        self.matrix = pd.read_excel(filepath, header=None).values
        index = self.find_sentinal_row('#')
        header = self.make_header(self.matrix[index, :], self.matrix[index+1, :])
        body = self.make_body(header, index)
        df =  pd.DataFrame(body, columns=header)
        #df['OFFICE NOTES: '] = df['OFFICE NOTES: '].apply(str) #might be useless
        self.data = df
        
    def find_sentinal_row(self, sentinal):
        i = 0
        row = self.matrix[i, :]
        while row[0] != sentinal:
            i = i + 1
            row = self.matrix[i, :]
        return i

    def make_header(self, r1, r2):
        r1 = np.array(r1, dtype='str')
        r2 = np.array(r2, dtype='str')
        last = len(r1)
        header = np.copy(r1)
        for i in range(len(r2)):
            if header[i] == 'nan' or header[i] == 'RESULT':
                header[i] = r2[i]
            if header[i] == 'OFFICE NOTES: ':
                last = i
        return header[0:last+1]

    def make_body(self, header, i):
        body = self.matrix[i+2:self.matrix.shape[0], 0:len(header)]
        return body


    def find_event_date(self):
        pass