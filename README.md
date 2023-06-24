# GraphGroupUnfurl

## Intro
This function app will:
 - Get a list of all non-dynamic security groups in AAD
 - Create and/or update groups that have fully expanded membership (membership is flattened from nested groups).

New groups are created and updated with UNF: prepended to the name.