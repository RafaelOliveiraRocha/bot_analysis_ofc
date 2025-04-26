#!/bin/bash
sudo apt-get update
sudo apt-get install -y locales
sudo locale-gen pt_BR.UTF-8
sudo update-locale LANG=pt_BR.UTF-8
