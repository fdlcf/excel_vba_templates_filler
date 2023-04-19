
import func
import path


table = func.open_table(path.table)
for index, row in table.iterrows():
    dic = {}
    agr_no = row['agr_no']
    print(agr_no)
    vehicle_cost = row['vehicle_cost']
    file_name = row['file_name']
    age = row['age']
    rv_net = row['rv_net']
    smr_diff = row['smr_diff']
    tire_diff = row['tire_diff']
    tire_storage = row['tire_storage']
    tire_change = row['tire_change']
    pzu = row['pzu']
    registration = row['registration']
    rsa = row['rsa']
    telem_install = row['telem_install']
    telem_service = row['telem_service']
    kasko = row['kasko']
    osago = row['osago']
    dsago = row['dsago']
    tax = row['tax']
    margin = row['margin']
    ets = row['ets']
    return_fee = row['return_fee']
    credit_remain = row['credit_remain']
    dic = {
        'agr_no': agr_no,
        'vehicle_cost': vehicle_cost,
        'file_name': file_name,
        'age': age,
        'rv_net': rv_net,
        'smr_diff': smr_diff,
        'tire_diff': tire_diff,
        'tire_storage': tire_storage,
        'tire_change': tire_change,
        'pzu': pzu,
        'registration': registration,
        'rsa': rsa,
        'telem_install': telem_install,
        'telem_service': telem_service,
        'kasko': kasko,
        'osago': osago,
        'dsago': dsago,
        'tax': tax,
        'margin': margin,
        'ets': ets,
        'return_fee': return_fee,
        'credit_remain': credit_remain,
    }
    func.fill_payment(dic)


