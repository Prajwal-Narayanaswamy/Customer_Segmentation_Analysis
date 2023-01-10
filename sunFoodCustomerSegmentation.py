from importer import *
from importer.fileImport import FileImporter
from importer.groupStatistics import GroupStatistics, stringifyGroup

defaultDataTypes = {
    'isMarried': int,
    'isEmployed': int,
    'hasNewBaby': int,
    'age': int,
    'annualIncome': float,
    'childrenNum': int,
    'avgPurchaseAmount': float
}

fileImporter = FileImporter(f'{DATA_FILE_PATH}SunFoodShop_customers.csv', defaultDataTypes=defaultDataTypes)
groupedData = fileImporter.getGroupData([
    ('hasNewBaby',),
    ('sex',),
    ('isMarried', 'isEmployed'),
    ('educationLevel', 'occupationCategory')
])

hasNewBabySegments = groupedData[('hasNewBaby',)]
babySegmentsStats = {groupKey: GroupStatistics(group) for groupKey, group in hasNewBabySegments.items()}
babySegmentHeaders = ['group', 'groupPct', 'count', 'avgPurchasePrice', 'maxPurchasePrice', 'minPurchasePrice', 'pctEmployed', 'avgAge', 'avgAnnualIncome']
babyAnalysisData = []

for groupKey, groupStats in babySegmentsStats.items():
    stats = groupStats.calculatedStatistics
    newRecord = []
    newRecord.append(stringifyGroup(('hasNewBaby',), groupKey))  # Implement the stringifyGroup function
    newRecord.append((stats['customerKey'][groupStats.COUNT] / len(fileImporter.data)) * 100)
    newRecord.append(stats['customerKey'][groupStats.COUNT])
    newRecord.append(round(stats['avgPurchaseAmount'][groupStats.MEAN], 2))
    newRecord.append(round(stats['avgPurchaseAmount'][groupStats.MAX], 2))
    newRecord.append(round(stats['avgPurchaseAmount'][groupStats.MIN], 2))
    newRecord.append(stats['isEmployed'][groupStats.MEAN] * 100)
    newRecord.append(stats['age'][groupStats.MEAN])
    newRecord.append(round(stats['annualIncome'][groupStats.MEAN], 2))
    babyAnalysisData.append(newRecord)

# Loop through each of the groupings in groupedData and create instances of GroupStatistics for each
# groupKey within each grouping
# Output should look something like {
#   ('hasNewBaby'): {(0,): <GroupStatistics instance>, (1,): <GroupStatistics instance>},
#   ('educationLevel', 'isMarried'): {('High school', 1): <GroupStatistics instance>,...},...
#  }
groupedSegmentStats = {}
for groupedDataKey, groupedSegments in groupedData.items():
    groupedSegmentStats[groupedDataKey] = {groupKey: GroupStatistics(group) for groupKey, group in groupedSegments.items()}

segmentationHeaders = ['group', 'customerCount', 'avgPurchasePrice']
segmentationAnalysisData = []
# Loop through groupedSegmentStats to format each GroupStatistics instance into the appropriate
# values based on the segmentationHeaders. Use the stringifyGroup function for the group value
# Output should look something like [
#   ['hasNewBaby: 0', 29, 403.27],
#   ['educationLevel: High school, isMarried: 1', 102, 209.80],...
# ]
for groupedDataKey, segmentStats in groupedSegmentStats.items():
    for groupKey, groupStats in segmentStats.items():
        newRecord = []
        stats = groupStats.calculatedStatistics
        newRecord.append(stringifyGroup(groupedDataKey, groupKey))
        newRecord.append(stats['customerKey'][groupStats.COUNT])
        newRecord.append(round(stats['avgPurchaseAmount'][groupStats.MEAN], 2))
        segmentationAnalysisData.append(newRecord)

sheetsConfig = [
    {'data': babyAnalysisData, 'headers': babySegmentHeaders, 'title': 'babyAnalysis'},
    {'data': segmentationAnalysisData, 'headers': segmentationHeaders, 'title': 'segmentationAnalysis'},
    {'title': 'rawData'},
]
fileImporter.writeExcelFile('sunFoodBabyAnalysis', sheetsConfig=sheetsConfig)
print('Finished')