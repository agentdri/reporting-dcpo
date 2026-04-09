#!/bin/bash
# Fix auto-generated service files: replace @microsoft/power-apps/data with @pa-client/power-code-sdk

SERVICES_DIR="src/generated/services"

if [ ! -d "$SERVICES_DIR" ]; then
  echo "Dossier $SERVICES_DIR introuvable."
  exit 1
fi

count=0
for file in "$SERVICES_DIR"/*Service.ts; do
  [ -f "$file" ] || continue

  if grep -q "@microsoft/power-apps/data" "$file"; then
    sed -i "s|import type { IOperationResult } from '@microsoft/power-apps/data';|import type { IOperationResult } from '@pa-client/power-code-sdk';|g" "$file"
    sed -i "s|import { getClient } from '@microsoft/power-apps/data';|import { getPowerSdkInstance } from '@pa-client/power-code-sdk';|g" "$file"
    sed -i "s|getClient(dataSourcesInfo)|getPowerSdkInstance(dataSourcesInfo).Data|g" "$file"
    echo "Fixed: $file"
    count=$((count + 1))
  fi
done

echo "Done. $count file(s) fixed."