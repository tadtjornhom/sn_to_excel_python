#!/bin/bash
if ps -ef | grep -v grep | grep "python main.py $sn_env" ; then
	exit 0	
else
	python main.py $sn_env
	exit 0
fi
