from pathlib import Path
from kep.output import save_all
from kep.reader import read_df
from kep.settings import Variable


def test_save_all(tempdir):
    xlpath = str(Path(__file__).parent.joinpath('agro.xlsx'))
    vs = [
            Variable('1.3.', 'B7', 'A', 'AGROPROD'),
            Variable('1.3.', 'C7', 'Q', 'AGROPROD'),
            Variable('1.3.', 'G7', 'A', 'AGROPROD')
        ]         
    save_all(xlpath, vs, tempdir        
    #FIXME save_all(xlpath, vs, tempdir) vs read_df() not working.    
    #assert list(save_all(xlpath, vs, tempdir)) == 1
    #dfa = read_df('a', tempdir)
    #assert dfa['1999'] ==  3 
         
    
    