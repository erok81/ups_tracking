import pickle

# Credentials
license_number = '0D7AEBD36E37D0FD'
user_id = 'tacomatrd2016'
password = '2003Yamaha'
# Email
email = 'erik.ottley@wdc.com'

config = {'license_number': '0D7AEBD36E37D0FD', 'user_id': 'tacomatrd2016', 'password': '2003Yamaha', 'email': 'erik.ottley@wdc.com'}


with open('config.pkl', 'wb') as f:
    pickle.dump(config, f)