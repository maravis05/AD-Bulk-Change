# AD-Bulk-Change
Takes input from the user (sAMAccountname) to populate a CSV template file with users to be changed. Once changes are made, user can then import a CSV file to make bulk changes to local Active Directory users.

Currently works with Title, Department and Manager. (headers can be added to change additional attributes)
Translates name to DistinguishedName for managers, so you only need first name and last name instead of full DN.
Does not make changes if the attribute is the same or if you leave the column blank in the CSV.
