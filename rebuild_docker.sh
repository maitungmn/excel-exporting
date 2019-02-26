#!/bin/bash
imageName=exportExcel:latest
containerName=exportExcel1
build=build

if [ $1 == $build ]
then

echo Delete old image ...
docker rmi -f $imageName
echo Build new Image
docker build -t $imageName -f Dockerfile .

else

echo Delete old image ...
docker rmi -f $imageName
echo Build new Image
docker build -t $imageName -f Dockerfile .
echo Delete old container...
docker rm -f $containerName
echo Run new container...
docker run -d -p 8080:8080 --name $containerName $imageName

fi
