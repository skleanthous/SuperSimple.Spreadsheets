Feature: CreationFromList
	In order to send a report to business
	As a developer
	I want to be able to create a spreadsheet from a list of data

@mytag
Scenario: Create a spreadsheet document from a list of data
	Given a list containing 5 items of a specific type
	When I call SaveToStream with the available data
	Then the result should be a file that can be opened
	And it should contain 6 rows (5 for items and one for header)
	And each row should correspond to the data in the list
