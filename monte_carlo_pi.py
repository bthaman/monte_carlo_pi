import numpy as np
import xlwings as xw
import math
import timeit


def convert_seconds(seconds):
    h = math.floor(seconds / 3600)
    str_h = '0' + str(h) if len(str(h)) < 2 else str(h)
    m = math.floor((seconds - h * 3600) / 60)
    str_m = '0' + str(m) if len(str(m)) < 2 else str(m)
    sec = round(seconds - h * 3600 - m * 60, 2)
    str_sec = '0' + str(sec) if str(sec).index('.') == 1 else str(sec)
    return str_h, str_m, str_sec


def xl_mc_pi():
    start_time = timeit.default_timer()  # time the execution
    sht = xw.Book.caller().sheets[0]
    sht.range('H6').value = ''
    # user Inputs
    num_trials = sht.range('E3').options(numbers=int).value
    chunks = sht.range('E4').options(numbers=int).value
    total_inside = 0
    estimates = np.zeros((chunks, 1))
    animate = sht.range('E9').value.lower() == 'yes'

    sht.range('O2').expand().clear_contents()
    sht.charts['Chart 5'].set_source_data(sht.range((1, 15), (chunks + 2, 17)))
    sht.range('O2').value = np.round(np.linspace(num_trials, num_trials * chunks, chunks).reshape(-1, 1), 2)
    sht.range('Q2').value = np.ones((chunks, 1)) * math.pi

    # have to run chunks of trials to avoid fatal barfing (running out of memory)
    # keep at or under 1e7
    for i in range(0, chunks):
        rands = np.random.random(size=(num_trials, 2))
        lst_vals = np.hypot(rands[:, 0], rands[:, 1])
        total_inside += int((lst_vals <= 1).sum())  # casting to int avoids overflow that occurs with np.int32 data type
        estimates[i, :] = 4 * total_inside / (num_trials * (i + 1))
        if animate:
            sht.range(i + 2, 16).value = estimates[i, 0]
            sht.range('E6').value = num_trials * (i + 1)
            sht.range('E7').value = estimates[i, 0]
            elapsed = timeit.default_timer() - start_time
            h, m, s = convert_seconds(elapsed)
            sht.range('H6').value = 'Execution time = {}:{}:{}'.format(h, m, s)
            sht.book.app.screen_updating = True
    if not animate:
        sht.range('P2').value = estimates
        sht.range('E6').value = num_trials * chunks
        sht.range('E7').value = estimates[-1, 0]
        elapsed = timeit.default_timer() - start_time
        h, m, s = convert_seconds(elapsed)
        sht.range('H6').value = 'Execution time = {}:{}:{}'.format(h, m, s)

if __name__ == "__main__":
    xl_mc_pi()
