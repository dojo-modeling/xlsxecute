FROM ubuntu:latest

RUN apt update && apt install -y python3-pip
RUN pip install formulas openpyxl

COPY run-excel-model /usr/local/bin/

