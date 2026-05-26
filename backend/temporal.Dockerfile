ARG TEMPORAL_VERSION=1.22.2

FROM temporalio/auto-setup:${TEMPORAL_VERSION}

COPY temporal/dynamicconfig /etc/temporal/config/dynamicconfig
