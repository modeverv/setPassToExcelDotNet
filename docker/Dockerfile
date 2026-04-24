FROM accetto/ubuntu-vnc-xfce-g3:latest

USER root
ENV DEBIAN_FRONTEND=noninteractive

RUN apt-get update && \
    apt-get install -y --no-install-recommends \
      libreoffice \
      libreoffice-l10n-ja \
      fonts-noto-cjk && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

WORKDIR /workspace
