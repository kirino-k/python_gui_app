#!/bin/bash

DIR_CMD=$(cd $(dirname $0); pwd)
DIR_ROOT=$(dirname ${DIR_CMD})
DIR_SRC=${DIR_ROOT}/src

# create exe file in docker container
docker run \
  --rm \
  --volume ${DIR_SRC}:/src \
  --volume /tmp/.X11-unix:/tmp/.X11-unix \
  --env DISPLAY=$DISPLAY \
  --entrypoint /bin/sh \
  account_management \
  -c "python main.py"

