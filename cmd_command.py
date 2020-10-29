# coding=utf8

from subprocess import check_output

def wincmd(my_command):
    val= (check_output(my_command, shell=True))
    return val
