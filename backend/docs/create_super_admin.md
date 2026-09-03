# Create super admin

```pwershell
$env:DATABASE_URL = "sqlite:///bootstrap_test.db"

uv run --no-sync python -m scripts.bootstrap_super_admin --email "bootstrap-admin@example.com" --full-name "Bootstrap Administrator"
```
