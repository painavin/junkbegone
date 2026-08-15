# Junk Be Gone — Azure Function

## Publish

```
cd az-function
func azure functionapp publish junkbegone
```

## Trigger `dailyCleanup` on demand

The Flex Consumption app's hostname includes a random suffix, so look it up dynamically rather than hardcoding it, along with the site's master key:

```
HOST=$(az functionapp show --name junkbegone --resource-group rg-sub-free --query defaultHostName -o tsv)
MASTER_KEY=$(az functionapp keys list --name junkbegone --resource-group rg-sub-free --query masterKey -o tsv)
curl -X POST "https://$HOST/admin/functions/dailyCleanup?code=$MASTER_KEY" -H "Content-Type: application/json" -d '{"input": ""}'
```

Check the results in the portal's **Log stream**, or via `az webapp log tail --name junkbegone --resource-group rg-sub-free`.

## Download `junk-senders.json`

Option A — connection string, run from inside `az-function/` (reads it from `.env`):

```
CONN=$(grep "^AZURE_STORAGE_CONNECTION_STRING=" .env | cut -d= -f2-)
az storage blob download --connection-string "$CONN" --container-name junkbegone --name junk-senders.json --file ./junk-senders.json
```

Option B — Azure AD login (requires a Storage Blob Data role assigned on the account):

```
az storage blob download \
  --account-name satickers \
  --container-name junkbegone \
  --name junk-senders.json \
  --file ./junk-senders.json \
  --auth-mode login
```

## Upload `junk-senders.json`

Option A — connection string:

```
CONN=$(grep "^AZURE_STORAGE_CONNECTION_STRING=" .env | cut -d= -f2-)
az storage blob upload --connection-string "$CONN" --container-name junkbegone --name junk-senders.json --file ./junk-senders.json --overwrite
```

Option B — Azure AD login (requires a Storage Blob Data role assigned on the account):

```
az storage blob upload \
  --account-name satickers \
  --container-name junkbegone \
  --name junk-senders.json \
  --file ./junk-senders.json \
  --overwrite \
  --auth-mode login
```
