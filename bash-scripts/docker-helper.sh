#login to azure via cli
az acr login --name regNimOsSBA

#lates/uat
docker buildx build . --platform linux/amd64 -t regnimossba.azurecr.io/regnimosssba/karma-mail:latest
docker push regnimossba.azurecr.io/regnimosssba/karma-mail:latest

#prod
docker buildx build . --platform linux/amd64 -t regnimossba.azurecr.io/regnimosssba/karma-mail:prod
docker push regnimossba.azurecr.io/regnimosssba/karma-mail:prod

#lis to confirm 
az acr repository list --name regNimOsSBA --output table

# run locally
docker run -p 80:3000 regnimossba.azurecr.io/regnimosssba/karma-mail:latest        


