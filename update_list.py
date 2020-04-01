import pickle

rma = input('Enter new RMA: ')
tracking = input('Enter new Tracking: ')


with open('tracking.pkl', 'rb') as f:
    tracking_nums = pickle.load(f)
    tracking_nums[rma] = tracking
    
with open('tracking.pkl', 'wb') as f:    
    pickle.dump(tracking_nums, f)