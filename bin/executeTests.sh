#!/usr/bin/env bash
nosetests --config ../test/nose.cfg --nologcapture -a 'unit,!ignore' ../. $@
