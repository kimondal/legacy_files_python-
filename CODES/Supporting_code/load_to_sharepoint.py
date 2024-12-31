from azure.identity import ClientSecretCredential
from azure.keyvault.secrets import SecretClient
from azure.core.exceptions import ResourceNotFoundError

# Azure AD application credentials
tenant_id = "your-tenant-id"
client_id = "your-client-id"
client_secret = "your-client-secret"

# Azure Key Vault URI
keyvault_uri = "https://your-keyvault-name.vault.azure.net/"

# Set up the credential (ClientSecretCredential for app-based authentication)
credential = ClientSecretCredential(tenant_id, client_id, client_secret)

# Create the SecretClient instance
secret_client = SecretClient(vault_url=keyvault_uri, credential=credential)


# Read all secrets in the Key Vault
def get_all_secrets():
    secret_list = []
    try:
        # List secrets in the Key Vault
        secrets = secret_client.list_properties_of_secrets()

        for secret in secrets:
            secret_name = secret.name
            secret_value = secret_client.get_secret(secret_name).value
            secret_list.append((secret_name, secret_value))
            print(f"Secret Name: {secret_name}, Secret Value: {secret_value}")

    except ResourceNotFoundError:
        print("Key Vault not found or no secrets available.")
    except Exception as e:
        print(f"An error occurred: {e}")

    return secret_list


# Call the function to read secrets
secrets = get_all_secrets()

# You can store the secrets in a list or process them as needed
print("\nAll secrets retrieved:")
for secret in secrets:
    print(f"Name: {secret[0]}, Value: {secret[1]}")
