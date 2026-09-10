"""Field-name catalog from Wildbook's public Bulk Import documentation.

Source: https://wildbook.docs.wildme.org/data/bulk-import-beta.html
Reviewed 2026-09-10. Indexed families use index 0 as an editable example.
The destination server may support a different subset or custom fields.
"""
FIELDS = sorted(set('''
MarkedIndividual.individualID Encounter.individualID Encounter.verbatimLocality
Encounter.locationID Encounter.decimalLatitude Encounter.decimalLongitude
Encounter.year Encounter.month Encounter.day Encounter.mediaAsset0 Encounter.genus
Encounter.specificEpithet Encounter.submitterID Encounter.state Encounter.alternateID
Encounter.behavior Encounter.country Encounter.dateInMilliseconds Encounter.distinguishingScar
Encounter.groupRole Encounter.hour Encounter.keyword0 Encounter.mediaAsset0.keywords
Encounter.lifeStage Encounter.livingStatus Encounter.measurement0 Encounter.mediaAsset0.[labelName]
Encounter.minutes Encounter.sightingID Encounter.researcherComments Encounter.sightingRemarks
Encounter.otherCatalogNumbers Encounter.patterningCode Encounter.mediaAsset0.quality Encounter.sex
MarkedIndividual.nickname Membership.role SatelliteTag.serialNumber SocialUnit.socialUnitName
Sighting.sightingID Sighting.comments Sighting.bestGroupSizeEstimate Sighting.effortCode
Sighting.fieldStudySite Sighting.fieldSurveyCode Sighting.groupBehavior Sighting.groupComposition
Sighting.groupSize Sighting.humanActivityNearby Sighting.individualCount Sighting.initialCue
Sighting.maxGroupSizeEstimate Sighting.millis Sighting.minGroupSizeEstimate Sighting.numAdults
Sighting.numAdultFemales Sighting.numAdultMales Sighting.numCalves Sighting.numJuveniles
Sighting.numSubAdults Sighting.numSubFemales Sighting.numSubMales Sighting.observer
Sighting.transectName Sighting.visibilityIndex Encounter.project0.projectIdPrefix
Encounter.project0.researchProjectName Encounter.project0.ownerUsername
Encounter.informOther0.affiliation Encounter.informOther0.emailAddress Encounter.informOther0.fullName
Encounter.photographer0.affiliation Encounter.photographer0.emailAddress Encounter.photographer0.fullName
Encounter.submitter0.affiliation Encounter.submitter0.emailAddress Encounter.submitter0.fullName
TissueSample.sampleID SexAnalysis.sex MitochondrialDNAAnalysis.haplotype
MicrosatelliteMarkersAnalysis.alleleNames MicrosatelliteMarkersAnalysis.alleles0
MicrosatelliteMarkersAnalysis.alleles1 SurveyTrack.vesselID survey.vessel Survey.id
Encounter.depth Encounter.Salinity Encounter.Salinity.samplingProtocol Sighting.bearing
Sighting.distance Sighting.seaState Sighting.seaSurfaceTemp Sighting.swellHeight Sighting.transectBearing
'''.split()))
