import sys
import time


def progressbar(it, prefix="", size=60, file=sys.stdout):
    count = len(it)

    def show(j):
        x = int(size * j / count)
        file.write("%s[%s%s] \r" % (prefix, "#" * x, "." * (size - x)))
        file.flush()

    show(0)
    for i, item in enumerate(it):
        yield item
        show(i + 1)
    file.write("\n")
    file.flush()


def wait(delay_in_sec):
    for i in progressbar(range(delay_in_sec), "\tWAIT: ", 40):
        time.sleep(1)