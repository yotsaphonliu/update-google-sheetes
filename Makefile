.PHONY: gcloud-login run-with-gcloud gcloud-all

gcloud-login:
	./scripts/gcloud_login.sh

excel-lookup:
	./scripts/excel_lookup.sh

run-with-gcloud:
	./scripts/run_with_gcloud.sh

gcloud-all: gcloud-login excel-lookup run-with-gcloud
