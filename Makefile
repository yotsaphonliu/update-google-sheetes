.PHONY: gcloud-login run-with-gcloud

gcloud-login:
	./scripts/gcloud_login.sh

run-with-gcloud:
	./scripts/run_with_gcloud.sh
