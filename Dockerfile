FROM ubuntu:latest
RUN apt-get update
RUN apt-get upgrade -y
RUN apt-get install git -y
RUN apt-get update
RUN apt-get install python3.8 -y --fix-missing
RUN apt install firefox -y
RUN apt install python3-pip -y

RUN git clone https://github.com/xWasp97x/Prenotazioni.git
WORKDIR ./Prenotazioni
CMD /bin/sh entrypoint.sh