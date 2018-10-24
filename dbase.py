import os, time
import pandas as pd
import pickle
from datetime import datetime
from dateutil.parser import parse


class Dbase:
    """ Class for storing Excel files in pickle object plus
    some maintenace utilities
    """

    def __init__(self, path, skiprows=11, nyears=60,
                 indexfile='pensioenfonds_index.xlsx'):
        self.path = path
        self.pickle_file = 'db.pkl'
        self.skiprows = skiprows
        self.nyears = nyears
        self.indexfile = indexfile

    def load(self):
        """ Load all Excel workbooks as given in pensioenfonds index
        """
        big_table = []
        filenames = self.read_index()
        for i, row in filenames.iterrows():
            filename, label, memo = row
            print('Loading {} (= {})'.format(filename, memo))
            infile = os.path.join(self.path, filename)
            df = self.read_xlswb(infile)
            multiindex = self.index2multiindex(df.columns)
            df.columns = multiindex
            df = self.stack(df, label)
            big_table.append(df)
        print("OK")
        return pd.concat(big_table, sort=False)

    def read_index(self):
        """ Read indexfile
        """
        infile = os.path.join(self.path, self.indexfile)
        return pd.read_excel(infile)

    def read_xlswb(self, xlswb):
        """ Load 2 sheets into 1 table
        """
        infile = os.path.join(self.path, xlswb)
        sheets = pd.read_excel(infile, sheet_name=[0, 1],
                               index_col=0, skiprows=self.skiprows,
                               usecols=2 * self.nyears)
        return pd.concat([sheets[0], sheets[1]], keys=['in', 'ac'],
                         names=['status'])

    def index2multiindex(self, colnames):
        """ Convert colnames to muliindex
        """
        newlabels = [int(str(lab).split('.')[0]) for lab in colnames]
        inflatie = self.nyears * ['no'] + self.nyears * ['re']
        tuples = list(zip(newlabels, inflatie))
        return pd.MultiIndex.from_tuples(tuples,
                                         names=['jaar', 'inflatie'])

    def stack(self, df, file_label):
        """ Stack to 1 large table
        """
        df = df.stack()
        df['file_label'] = file_label
        df.reset_index(inplace=True)
        rowindex = ['file_label', 'status', 'inflatie', 'scenario']
        return df.set_index(rowindex)

    def dump(self, obj):
        """ Dumps given object to pickle file
        """
        # dump object to pickle file
        pickle_file = os.path.join(self.path, self.pickle_file)
        f = open(pickle_file, 'wb')
        pickle.dump(obj,f)
        f.close()

    def get_pickle(self):
        return os.path.join(self.path, self.pickle_file)

    def OLD_load(self):
        fname = self.get_pickle()
        f = open(fname, 'rb')
        df = pickle.load(f)
        f.close()
        return df

    def get_hours_since_last_modified(self, filename):
        """ Return number of hours since last modifation of given file
        """
        last_modified = time.ctime(os.path.getmtime(filename))
        last_mod = parse(last_modified)
        now = datetime.now()
        ago = (now - last_mod)
        hours_ago = ago.seconds / (60 * 60)
        return hours_ago

    def update(self):
        files = self.read()
        _ = self.dump(files)
