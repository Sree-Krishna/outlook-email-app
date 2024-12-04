from azure.identity import AuthorizationCodeCredential, TokenCachePersistenceOptions

from msgraph import GraphServiceClient
from config import Config
from flask import g

def get_credential(auth_code=None):
    """
    Initialize and return the Microsoft Graph client.
    """
    config = Config()
    cache_options = TokenCachePersistenceOptions(name="graph_token_cache")
    credential = AuthorizationCodeCredential(
        tenant_id=config.TENANT_ID,
        client_id=config.CLIENT_ID,
        client_secret=config.CLIENT_SECRET,
        authorization_code=auth_code,
        redirect_uri=config.REDIRECT_URI,
        cache_persistence_options=cache_options
    )
    return credential

def get_graph_client(auth_code=None):
    credential = get_credential(auth_code)
    return GraphServiceClient(credential)
