"""


"""


import sys
from Mazama_CountryDictionaries import *

# define some exceptions for this package
class CountryTranslationError(Exception): pass

class CountryTranslator:
    dialects = ['BP_2011','BP_2012', 'BP_2015']

    def __init__(self,dialect='BP_2012'):
        if dialect in self.dialects:
            self.dialect = dialect
        else:
            sys.exit(1)

    def get_MZM_code(self,country):
        ########################################
        # get_MZM_code
        #
        # Converts BP country names to ISO-3166 two-character country codes.
        #
        # ISO codes obtained from:
        # http://www.iso.org/iso/en/prods-services/iso3166ma/02iso-3166-code-lists/list-en1-semic.txt

        # Convert country argument to upper case
        up_country = country.strip().upper()

        try: 
            # Attempt to convert all region names to ISO codes.
            code = English_to_ISO[up_country]
        except:
            # Catch any exceptions and test those against other acceptable names, depending on the dialect.
            try:
                if self.dialect == 'BP_2011' or self.dialect == 'BP_2015':
                    code = BP_to_MZM[up_country]
            except:
                # Regions that are still unrecognized
                error_string = "Cannot convert \"%s\" to ISO code" % (up_country)
                raise CountryTranslationError, error_string
                ###sys.exit(1)

        return code


