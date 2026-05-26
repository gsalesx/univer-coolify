ARG EDITION=release
ARG UNIVERSER_VERSION=0.24.0

FROM univer-acr-registry.cn-shenzhen.cr.aliyuncs.com/${EDITION}/universer:${UNIVERSER_VERSION}

COPY configs /data/configs
