using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace Daniel.CrmExcel
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Linq;
    using System.Windows.Forms;
    using System.Xml;

    using Ionic.Zip;


    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;

    using OfficeOpenXml;

    public class CrmToExcel : PluginControlBase
    {
        private readonly List<string> exclusionAttributes = new List<string>();

        private XmlDocument configurationFile;

        private readonly Dictionary<string, EntityMetadata> dicEntities = new Dictionary<string, EntityMetadata>();

        private ExcelWorksheet worksheetAttributes;

        private ExcelWorksheet worksheetEntities;

        private ExcelWorksheet worksheetRelationShipsManyToMany;

        private ExcelWorksheet worksheetRelationShipsManyToOne;

        public int LanguageCode
        {
            get;
            private set;
        }

        public CrmToExcel(IOrganizationService service, BackgroundWorker backgroundWorker)
        {
            this.Service = service;
            this.BackgroundWorker = backgroundWorker;
            this.exclusionAttributes.Clear();
            this.exclusionAttributes.Add("modifiedby");
            this.exclusionAttributes.Add("modifiedon");
            this.exclusionAttributes.Add("createdby");
            this.exclusionAttributes.Add("createdon");
            this.exclusionAttributes.Add("createdonbehalfby");
            this.exclusionAttributes.Add("ownerid");
            this.exclusionAttributes.Add("ownershipcode");
            this.exclusionAttributes.Add("owningbusinessunit");
            this.exclusionAttributes.Add("owningteam");
            this.exclusionAttributes.Add("owninguser");
            this.exclusionAttributes.Add("modifiedonbehalfby");
            this.exclusionAttributes.Add("importsequencenumber");
            this.exclusionAttributes.Add("overriddencreatedon");
            this.exclusionAttributes.Add("timezoneruleversionnumber");
            this.exclusionAttributes.Add("utcconversiontimezonecode");
            this.exclusionAttributes.Add("versionnumber");
            //this.exclusionAttributes.Add("statecode");
            //this.exclusionAttributes.Add("statuscode");
        }

        public bool UseSolutionXml
        {
            get;
            private set;
        }

        public bool IncludeOwnerInformation
        {
            get;
            private set;
        }

        public IOrganizationService Service
        {
            get;
            private set;
        }

        public BackgroundWorker BackgroundWorker
        {
            get;
        }

        internal void CrmToExcelSheet(string solution, string excelFile, List<string> selectedNodes, bool useSolutionXml, bool includeOwner, int languageCode)
        {
            this.LanguageCode = languageCode;
            this.UseSolutionXml = useSolutionXml;
            this.IncludeOwnerInformation = includeOwner;
            LogInfo("Start Export solution");

            if (this.UseSolutionXml)
            {
                this.RetrievingSolution(solution);
            }

            LogInfo("Start creating new Excel file");
            var crmExcel = new FileInfo(excelFile);
            if (crmExcel.Exists)
            {
                LogInfo("Delete existing excel file.");
                crmExcel.Delete();
            }

            crmExcel = new FileInfo(excelFile);
            this.BackgroundWorker.ReportProgress(10, "Creating Excel sheet");
            using (var excelPackage = new ExcelPackage(crmExcel))
            {
                // Get handle to the existing worksheet
                this.worksheetEntities = excelPackage.Workbook.Worksheets.Add("Entities");
                this.worksheetAttributes = excelPackage.Workbook.Worksheets.Add("Attributes");
                this.worksheetRelationShipsManyToOne = excelPackage.Workbook.Worksheets.Add("RelationShips Many to One");
                this.worksheetRelationShipsManyToMany = excelPackage.Workbook.Worksheets.Add("RelationShips Many to Many");

                if (this.worksheetEntities != null)
                {
                    this.AddColumnHeadings();

                    var rowEntityIndex = 2;
                    var rowAttributeIndex = 2;
                    var rowManyToOneRelationshipIndex = 2;
                    var rowManyToManyRelationshipIndex = 2;
                    foreach (var node in selectedNodes)
                    {
                        this.BackgroundWorker.ReportProgress(10, $"Adding entity '{node}' to Excel sheet");
                        var entityName = node;
                        LogInfo($"Add entity {entityName} to sheet.");
                        this.AddEntityToSheet(ref rowEntityIndex, ref rowAttributeIndex, ref rowManyToOneRelationshipIndex, ref rowManyToManyRelationshipIndex, entityName);
                    }
                }

               excelPackage.Save();
            }
        }

        private void RetrievingSolution(string solution)
        {
            this.BackgroundWorker.ReportProgress(10, $"Retrieving Solution '{solution}'");
            var exportSolutionRequest = new ExportSolutionRequest();
            exportSolutionRequest.Managed = false;
            exportSolutionRequest.SolutionName = solution; //e2.Argument.ToString();
            try
            {
                var exportSolutionResponse = (ExportSolutionResponse)this.Service.Execute(exportSolutionRequest);
                if (exportSolutionResponse.ExportSolutionFile.Length == 0)
                {
                    LogInfo("Customization file empty.");
                }
                else
                {
                    var exportXml = exportSolutionResponse.ExportSolutionFile;
                    var ms = new MemoryStream(exportXml);
                    LogInfo("Unzip customizations xml");
                    using (var zip = ZipFile.Read(ms))
                    {
                        foreach (var zippedFile in zip)
                        {
                            if (zippedFile.FileName == "customizations.xml")
                            {
                                zippedFile.Extract(Application.StartupPath, ExtractExistingFileAction.OverwriteSilently);
                            }
                        }
                    }
                    LogInfo("Customization file extracted.");
                    this.configurationFile = new XmlDocument();
                    this.configurationFile.Load(Path.Combine(Application.StartupPath, "customizations.xml"));
                }
            }
            catch (Exception exception)
            {
                this.BackgroundWorker.ReportProgress(10, $"Retrieving Solution Error - {exception.Message}");
                LogError("Error getting solution", exception);
            }
        }

        private void AddColumnHeadings()
        {
            this.worksheetEntities.Cell(1, Constants.ColSchemaName).Value = "SchemaName (C)";
            this.worksheetEntities.Cell(1, Constants.ColOwnerShipType).Value = "OwnerShipType (C)";
            this.worksheetEntities.Cell(1, Constants.ColEntityDisplayName).Value = "DisplayName";
            this.worksheetEntities.Cell(1, Constants.ColDisplayCollectionName).Value = "DisplayCollectionName (C)";
            this.worksheetEntities.Cell(1, Constants.ColDescription).Value = "Description (C)";
            this.worksheetEntities.Cell(1, Constants.ColPrimaryAttributeSchemaName).Value = "PrimaryAttributeSchemaName (C)";
            this.worksheetEntities.Cell(1, Constants.ColPrimaryAttributeDisplayname).Value = "PrimaryAttributeDisplayname (C)";
            this.worksheetEntities.Cell(1, Constants.ColPrimaryAttributeDescription).Value = "PrimaryAttributeDescription (C)";

            this.worksheetAttributes.Cell(1, Constants.ColEntityError).Value = "Error messages";
            this.worksheetEntities.Cell(1, Constants.ColIsActivity).Value = "As activity entity (C)";

            //Gets or sets whether a custom activity should appear in the activity menus in the Web application.
            this.worksheetEntities.Cell(1, Constants.ColActivityTypeMask).Value = "ActivityTypeMask (C)";

            //Business process flows
            this.worksheetEntities.Cell(1, Constants.ColIsBusinessProcessEnabled).Value = "IsBusinessProcessEnabled (C)";

            // Connections;
            this.worksheetEntities.Cell(1, Constants.ColIsConnectionsEnabled).Value = "IsConnectionsEnable (C)d";

            this.worksheetEntities.Cell(1, Constants.ColIsEmailEnabled).Value = "IsEmailEnabled (C)";
            // Mail merge
            this.worksheetEntities.Cell(1, Constants.ColIsMailMergeEnabled).Value = "IsMailMergeEnabled (C)";

            this.worksheetEntities.Cell(1, Constants.ColIsDocumentManagementEnabled).Value = "IsDocumentManagementEnabled (C)";

            this.worksheetEntities.Cell(1, Constants.ColAutoCreateAccessTeams).Value = "AutoCreateAccessTeams (C)";

            // Audit
            this.worksheetEntities.Cell(1, Constants.ColIsAuditEnabled).Value = "IsAuditEnabled (C)";

            this.worksheetEntities.Cell(1, Constants.ColAutoRouteToOwnerQueue).Value = "AutoRouteToOwnerQueue (C)";

            this.worksheetEntities.Cell(1, Constants.ColIsQuickCreateEnabled).Value = "IsQuickCreateEnabled (C)";

            this.worksheetEntities.Cell(1, Constants.ColPrimaryAttributeLength).Value = "PrimaryAttributeLength (C)";
            this.worksheetEntities.Cell(1, Constants.ColPrimaryAttributeRequired).Value = "PrimaryAttributeRequired (C)";
            this.worksheetEntities.Cell(1, Constants.ColEntityError).Value = "Errors";
            //worksheetEntities.Cell(1, colHasActivities).Value = "HasActivities";
            //worksheetEntities.Cell(1, colHasNotes).Value = "HasNotes";

            this.worksheetAttributes.Cell(1, 1).Value = "EntityLogicalName (C)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeType).Value = "Type (C)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeSchemaName).Value = "SchemaName (C)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeTab).Value = "Form-Tab-Section-Label (-)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeDisplayName).Value = "DisplayName (C,U)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeMaxLength).Value = "MaxLength (C,U)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeOptionset).Value = "Optionset values (C)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeMin).Value = "Min value (C)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeMax).Value = "Max value (C)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeIsAuditEnabled).Value = "IsAuditEnabled (C,U)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeIsValidForAdvancedFind).Value = "IsValidForAdvancedFind (C,U)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeRequired).Value = "Required (C,U)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeStringFormat).Value = "String Format (C)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeDescription).Value = "Description (C,U)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeForm).Value = "Form (-)";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeDateBehavior).Value = "Datebehaviour";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeDateFormat).Value = "DateFormat";
            this.worksheetAttributes.Cell(1, Constants.ColAttributeErrors).Value = "Error messages";

            this.worksheetRelationShipsManyToOne.Cell(1, 1).Value = "EntityLogicalName (C)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipSchemaName).Value = "SchemaName (C)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipReferencedEntity).Value = "ReferencedEntity-primary (C)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipReferencedAttribute).Value = "ReferencedAttribute (-)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipReferencingAttribute).Value = "ReferencingAttribute (-)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipReferencingEntity).Value = "ReferencingEntity (C)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipType).Value = "Relation Type (-)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipLabel).Value = "Label (seen on left hand wunderbar)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipLookUpSchemaName).Value = "LookUpSchemaName (C)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipLookUpDisplayName).Value = "LookUpDisplayName (C)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipLookUpRequiredLevel).Value = "LookUpRequiredLevel (C)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipLookUpDescription).Value = "LookUpDescription (C)";

            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipCascadeConfigurationAssign).Value = "Cascade.Assign (-)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipCascadeConfigurationDelete).Value = "Cascade.Delete (-)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipCascadeConfigurationMerge).Value = "Cascade.Merge (-)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipCascadeConfigurationReparent).Value = "Cascade.Reparent (-)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipCascadeConfigurationShare).Value = "Cascade.Share (-)";
            this.worksheetRelationShipsManyToOne.Cell(1, Constants.ColRelationShipCascadeConfigurationUnshare).Value = "Cascade.UnShare (-)";

            this.worksheetRelationShipsManyToMany.Cell(1, 1).Value = "EntityLogicalName";
            this.worksheetRelationShipsManyToMany.Cell(1, Constants.ColRelationShipMmSchemaName).Value = "SchemaName";
            this.worksheetRelationShipsManyToMany.Cell(1, Constants.ColRelationShipMmIntersectEntitySchemaName).Value = "Intersect Entity SchemaName";
            this.worksheetRelationShipsManyToMany.Cell(1, Constants.ColRelationShipMmEntity1IntersectAttribute).Value = "Entity 1 Intersect Attribute";
            this.worksheetRelationShipsManyToMany.Cell(1, Constants.ColRelationShipMmEntity1LogicalName).Value = "Entity 1 Logical Name";
            this.worksheetRelationShipsManyToMany.Cell(1, Constants.ColRelationShipMmEntity2IntersectAttribute).Value = "Entity 2 Intersect Attribute";
            this.worksheetRelationShipsManyToMany.Cell(1, Constants.ColRelationShipMmEntity2LogicalName).Value = "Entity 2 Logical Name";
            this.worksheetRelationShipsManyToMany.Cell(1, Constants.ColRelationShipMmEntity1AssociatedMenuLabel).Value = "Entity 1 Associated Menu Label";
            this.worksheetRelationShipsManyToMany.Cell(1, Constants.ColRelationShipMmEntity2AssociatedMenuLabel).Value = "Entity 2 Associated Menu Label";
        }

        public EntityMetadata RetrieveEntity(string logicalName)
        {
            var retrieveEntityRequest = new RetrieveEntityRequest
                                            {
                                                LogicalName = logicalName.ToLower(),
                                                EntityFilters = EntityFilters.All
                                            };
            try
            {
                var retrieveEntityResponse = (RetrieveEntityResponse)this.Service.Execute(retrieveEntityRequest);
                return retrieveEntityResponse.EntityMetadata;
            }
            catch (Exception exception)
            {
                LogError($"Error retrieving entityt metadata for entity '{logicalName}'", exception);
            }

            return null;
        }

        private void AddEntityToSheet(ref int rowEntityIndex, ref int rowAttributeIndex, ref int rowManyToOneRelationshipIndex, ref int rowManyToManyRelationshipIndex, string entityLogicalName)
        {
            var entityMetadata = this.RetrieveEntity(entityLogicalName);
            //https://msdn.microsoft.com/en-us/library/microsoft.xrm.sdk.metadata.entitymetadata_members.aspx
            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColSchemaName).Value = entityMetadata.SchemaName;
            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColOwnerShipType).Value = entityMetadata.OwnershipType.Description();
            //OwnershipTypes.UserOwned
            //OwnershipTypes.OrganizationOwned
            if (entityMetadata.DisplayName.LocalizedLabels.Count > 0)
            {
                this.worksheetEntities.Cell(rowEntityIndex, Constants.ColEntityDisplayName).Value = entityMetadata.DisplayName.UserLocalizedLabel.Label;
            }

            if (entityMetadata.DisplayCollectionName.LocalizedLabels.Count > 0)
            {
                if (entityMetadata.DisplayCollectionName.UserLocalizedLabel != null)
                {
                    this.worksheetEntities.Cell(rowEntityIndex, Constants.ColDisplayCollectionName).Value = entityMetadata.DisplayCollectionName.UserLocalizedLabel.Label;
                }
            }

            if (entityMetadata.Description.LocalizedLabels.Count > 0)
            {
                if (entityMetadata.Description.UserLocalizedLabel != null)
                {
                    this.worksheetEntities.Cell(rowEntityIndex, Constants.ColDescription).Value = entityMetadata.Description.UserLocalizedLabel.Label;
                }
            }

            // Gets or sets whether the entity is an activity.
            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColIsActivity).Value = entityMetadata.IsActivity.GetValueOrDefault(false).ToString();

            // Gets or sets whether a custom activity should appear in the activity menus in the Web application.
            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColActivityTypeMask).Value = entityMetadata.ActivityTypeMask.GetValueOrDefault(0).ToString();

            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColIsBusinessProcessEnabled).Value = entityMetadata.IsBusinessProcessEnabled.GetValueOrDefault(false).ToString();

            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColIsConnectionsEnabled).Value = entityMetadata.IsConnectionsEnabled.Value.ToString();

            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColIsMailMergeEnabled).Value = entityMetadata.IsMailMergeEnabled.Value.ToString();

            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColIsEmailEnabled).Value = entityMetadata.IsActivityParty.GetValueOrDefault(false).ToString();

            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColIsDocumentManagementEnabled).Value = entityMetadata.IsDocumentManagementEnabled.GetValueOrDefault(false).ToString();

            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColAutoCreateAccessTeams).Value = entityMetadata.AutoCreateAccessTeams.GetValueOrDefault(false).ToString();

            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColIsAuditEnabled).Value = entityMetadata.IsAuditEnabled.Value.ToString();

            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColAutoRouteToOwnerQueue).Value = entityMetadata.AutoRouteToOwnerQueue.GetValueOrDefault(false).ToString();

            this.worksheetEntities.Cell(rowEntityIndex, Constants.ColIsQuickCreateEnabled).Value = entityMetadata.IsQuickCreateEnabled.GetValueOrDefault(false).ToString();

            if (entityMetadata.PrimaryNameAttribute != null)
            {
                this.worksheetEntities.Cell(rowEntityIndex, Constants.ColPrimaryAttributeSchemaName).Value = entityMetadata.PrimaryNameAttribute;

                var primaryAttribute = entityMetadata.Attributes.FirstOrDefault(a => a.SchemaName == entityMetadata.PrimaryNameAttribute.ToLower());
                if (primaryAttribute != null)
                {
                    this.worksheetEntities.Cell(rowEntityIndex, Constants.ColPrimaryAttributeRequired).Value = primaryAttribute.RequiredLevel.Value.Description();
                    var stringAttributeMetadata = (StringAttributeMetadata)primaryAttribute;
                    this.worksheetEntities.Cell(rowEntityIndex, Constants.ColPrimaryAttributeLength).Value = stringAttributeMetadata.MaxLength.ToString();

                    if (primaryAttribute.DisplayName.LocalizedLabels.Count > 0)
                    {
                        if (primaryAttribute.DisplayName.UserLocalizedLabel != null)
                        {
                            if (primaryAttribute.DisplayName.UserLocalizedLabel.Label != null)
                            {
                                this.worksheetEntities.Cell(rowEntityIndex, Constants.ColPrimaryAttributeDisplayname).Value = primaryAttribute.DisplayName.UserLocalizedLabel.Label;
                            }
                        }
                    }

                    if (primaryAttribute.Description.LocalizedLabels.Count > 0)
                    {
                        if (primaryAttribute.Description.UserLocalizedLabel != null)
                        {
                            if (primaryAttribute.Description.UserLocalizedLabel.Label != null)
                            {
                                this.worksheetEntities.Cell(rowEntityIndex, Constants.ColPrimaryAttributeDescription).Value = primaryAttribute.Description.UserLocalizedLabel.Label;
                            }
                        }
                    }
                }
            }

            rowEntityIndex++;
            this.AddAttributesToExcelSheet(entityMetadata, ref rowAttributeIndex);

            this.AddManyToOneRelationshipsToSheet(ref rowManyToOneRelationshipIndex, entityLogicalName, entityMetadata);

            this.AddOneToManyRelationshipsToSheet(ref rowManyToOneRelationshipIndex, entityLogicalName, entityMetadata);

            this.AddManyToManyRelationshipsToSheet(ref rowManyToManyRelationshipIndex, entityLogicalName, entityMetadata);
        }

        private void AddManyToOneRelationshipsToSheet(ref int rowManyToOneRelationshipIndex, string entityLogicalName, EntityMetadata entityMetadata)
        {
            foreach (var manyToOneRelationship in entityMetadata.ManyToOneRelationships.OrderBy(a => a.SchemaName))
            {
                if (!manyToOneRelationship.IsCustomRelationship.GetValueOrDefault(false) && !this.IncludeOwnerInformation)
                {
                    continue;
                }

                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, 1).Value = entityLogicalName;
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipSchemaName).Value = manyToOneRelationship.SchemaName.ToLower();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipReferencedEntity).Value = manyToOneRelationship.ReferencedEntity.ToLower();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipReferencedAttribute).Value = manyToOneRelationship.ReferencedAttribute.ToLower();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipReferencingEntity).Value = manyToOneRelationship.ReferencingEntity.ToLower();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipReferencingAttribute).Value = manyToOneRelationship.ReferencingAttribute.ToLower();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipType).Value = "Many to One";

                // Get this data from the referencing enitity
                EntityMetadata entityMetaData;
                if (this.dicEntities.ContainsKey(manyToOneRelationship.ReferencingEntity))
                {
                    entityMetaData = this.dicEntities[manyToOneRelationship.ReferencingEntity];
                }
                else
                {
                    entityMetaData = this.RetrieveEntity(manyToOneRelationship.ReferencingEntity);
                    this.dicEntities.Add(manyToOneRelationship.ReferencingEntity, entityMetaData);
                }

                if (manyToOneRelationship.AssociatedMenuConfiguration?.Label?.UserLocalizedLabel != null)
                {
                    if (manyToOneRelationship.AssociatedMenuConfiguration.Label.UserLocalizedLabel.Label != null)
                    {
                        this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLabel).Value = manyToOneRelationship.AssociatedMenuConfiguration.Label.UserLocalizedLabel.Label;
                    }
                }

                var attribute = entityMetaData.Attributes.FirstOrDefault(a => string.Equals(a.SchemaName, manyToOneRelationship.ReferencingAttribute, StringComparison.CurrentCultureIgnoreCase));
                if (attribute != null)
                {
                    this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLookUpSchemaName).Value = attribute.SchemaName.ToLower();
                    if (attribute.DisplayName != null)
                    {
                        if (attribute.DisplayName.UserLocalizedLabel.Label != null)
                        {
                            this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLookUpDisplayName).Value = attribute.DisplayName.UserLocalizedLabel.Label;
                        }
                    }

                    if (attribute.Description?.UserLocalizedLabel != null)
                    {
                        if (attribute.Description.UserLocalizedLabel.Label != null)
                        {
                            this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLookUpDescription).Value = attribute.Description.UserLocalizedLabel.Label;
                        }
                    }

                    this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLookUpRequiredLevel).Value = attribute.RequiredLevel.Value.Description();
                }

                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationAssign).Value = manyToOneRelationship.CascadeConfiguration?.Assign?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationDelete).Value = manyToOneRelationship.CascadeConfiguration?.Delete?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationMerge).Value = manyToOneRelationship.CascadeConfiguration?.Merge?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationReparent).Value = manyToOneRelationship.CascadeConfiguration?.Reparent?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationShare).Value = manyToOneRelationship.CascadeConfiguration?.Share?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationUnshare).Value = manyToOneRelationship.CascadeConfiguration?.Unshare?.Description();

                rowManyToOneRelationshipIndex++;
            }
        }

        private void AddOneToManyRelationshipsToSheet(ref int rowManyToOneRelationshipIndex, string entityLogicalName, EntityMetadata entityMetadata)
        {
            foreach (var manyToOneRelationship in entityMetadata.OneToManyRelationships.OrderBy(a => a.SchemaName))
            {
                if (!manyToOneRelationship.IsCustomRelationship.GetValueOrDefault(false) && !this.IncludeOwnerInformation)
                {
                    continue;
                }

                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, 1).Value = entityLogicalName.ToLower();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipSchemaName).Value = manyToOneRelationship.SchemaName.ToLower();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipReferencedEntity).Value = manyToOneRelationship.ReferencedEntity;
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipReferencedAttribute).Value = manyToOneRelationship.ReferencedAttribute;
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipReferencingEntity).Value = manyToOneRelationship.ReferencingEntity;
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipReferencingAttribute).Value = manyToOneRelationship.ReferencingAttribute;
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipType).Value = "One to Many";

                if (manyToOneRelationship.AssociatedMenuConfiguration?.Label?.UserLocalizedLabel?.Label != null)
                {
                    this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLabel).Value = manyToOneRelationship.AssociatedMenuConfiguration?.Label?.UserLocalizedLabel?.Label;

                }

                EntityMetadata entityMetaData;
                if (this.dicEntities.ContainsKey(manyToOneRelationship.ReferencingEntity))
                {
                    entityMetaData = this.dicEntities[manyToOneRelationship.ReferencingEntity];
                }
                else
                {
                    entityMetaData = this.RetrieveEntity(manyToOneRelationship.ReferencingEntity);
                    this.dicEntities.Add(manyToOneRelationship.ReferencingEntity, entityMetaData);
                }

                var attribute = entityMetaData.Attributes.FirstOrDefault(a => string.Equals(a.SchemaName, manyToOneRelationship.ReferencingAttribute, StringComparison.CurrentCultureIgnoreCase));
                if (attribute != null)
                {
                    this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLookUpSchemaName).Value = attribute.SchemaName.ToLower();
                    if (attribute.DisplayName?.UserLocalizedLabel?.Label != null)
                    {
                        this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLookUpDisplayName).Value = attribute.DisplayName?.UserLocalizedLabel?.Label;

                        if (attribute.Description?.UserLocalizedLabel != null)
                        {
                            this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLookUpDescription).Value = attribute.Description?.UserLocalizedLabel?.Label;
                        }
                    }

                    this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipLookUpRequiredLevel).Value = attribute.RequiredLevel?.Value.Description();
                }


                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationAssign).Value = manyToOneRelationship.CascadeConfiguration?.Assign?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationDelete).Value = manyToOneRelationship.CascadeConfiguration?.Delete?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationMerge).Value = manyToOneRelationship.CascadeConfiguration?.Merge?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationReparent).Value = manyToOneRelationship.CascadeConfiguration?.Reparent?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationShare).Value = manyToOneRelationship.CascadeConfiguration?.Share?.Description();
                this.worksheetRelationShipsManyToOne.Cell(rowManyToOneRelationshipIndex, Constants.ColRelationShipCascadeConfigurationUnshare).Value = manyToOneRelationship.CascadeConfiguration?.Unshare?.Description();

                rowManyToOneRelationshipIndex++;
            }
        }

        private void AddManyToManyRelationshipsToSheet(ref int rowManyToManyRelationshipIndex, string entityLogicalName, EntityMetadata entityMetadata)
        {
            foreach (var manyToManyRelationship in entityMetadata.ManyToManyRelationships.OrderBy(a => a.SchemaName))
            {
                this.worksheetRelationShipsManyToMany.Cell(rowManyToManyRelationshipIndex, 1).Value = entityLogicalName.ToLower();
                this.worksheetRelationShipsManyToMany.Cell(rowManyToManyRelationshipIndex, Constants.ColRelationShipMmIntersectEntitySchemaName).Value = manyToManyRelationship.IntersectEntityName;
                this.worksheetRelationShipsManyToMany.Cell(rowManyToManyRelationshipIndex, Constants.ColRelationShipMmSchemaName).Value = manyToManyRelationship.SchemaName.ToLower();
                this.worksheetRelationShipsManyToMany.Cell(rowManyToManyRelationshipIndex, Constants.ColRelationShipMmEntity1IntersectAttribute).Value = manyToManyRelationship.Entity1IntersectAttribute;
                this.worksheetRelationShipsManyToMany.Cell(rowManyToManyRelationshipIndex, Constants.ColRelationShipMmEntity1LogicalName).Value = manyToManyRelationship.Entity1LogicalName;
                this.worksheetRelationShipsManyToMany.Cell(rowManyToManyRelationshipIndex, Constants.ColRelationShipMmEntity2IntersectAttribute).Value = manyToManyRelationship.Entity2IntersectAttribute;
                this.worksheetRelationShipsManyToMany.Cell(rowManyToManyRelationshipIndex, Constants.ColRelationShipMmEntity2LogicalName).Value = manyToManyRelationship.Entity2LogicalName;
                if (manyToManyRelationship.Entity1AssociatedMenuConfiguration?.Label?.UserLocalizedLabel?.Label != null)
                {
                    this.worksheetRelationShipsManyToMany.Cell(rowManyToManyRelationshipIndex, Constants.ColRelationShipMmEntity1AssociatedMenuLabel).Value = manyToManyRelationship.Entity1AssociatedMenuConfiguration?.Label?.UserLocalizedLabel?.Label;
                    this.worksheetRelationShipsManyToMany.Cell(rowManyToManyRelationshipIndex, Constants.ColRelationShipMmEntity2AssociatedMenuLabel).Value = manyToManyRelationship.Entity1AssociatedMenuConfiguration?.Label?.UserLocalizedLabel?.Label;
                }

                rowManyToManyRelationshipIndex++;
            }
        }

        private void AddAttributesToExcelSheet(EntityMetadata entityMetadata, ref int rowAttributeIndex)
        {
            foreach (var attributeMetadata in entityMetadata.Attributes.OrderBy(a => a.LogicalName))
            {
                if (attributeMetadata.AttributeType != AttributeTypeCode.Virtual && string.IsNullOrEmpty(attributeMetadata.AttributeOf) && !this.exclusionAttributes.Contains(attributeMetadata.SchemaName.ToLower()))
                {
                    this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColSchemaName).Value = entityMetadata.SchemaName;

                    this.AddAttributeFormInformationToSheet(entityMetadata, rowAttributeIndex, attributeMetadata);

                    this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeType).Value = attributeMetadata.AttributeType?.Description();
                    this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeSchemaName).Value = attributeMetadata.SchemaName;
                    this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeRequired).Value = attributeMetadata.RequiredLevel.Value.Description();

                    this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeIsAuditEnabled).Value = attributeMetadata.IsAuditEnabled.Value.ToString();
                    this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeIsValidForAdvancedFind).Value = attributeMetadata.IsValidForAdvancedFind.Value.ToString();

                    if (attributeMetadata.DisplayName.LocalizedLabels.Count > 0)
                    {
                        if (attributeMetadata.DisplayName.UserLocalizedLabel != null)
                        {
                            if (attributeMetadata.DisplayName.UserLocalizedLabel.Label != null)
                            {
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeDisplayName).Value = attributeMetadata.DisplayName.UserLocalizedLabel.Label;
                            }
                        }
                    }

                    if (attributeMetadata.Description.LocalizedLabels.Count > 0)
                    {
                        if (attributeMetadata.Description.UserLocalizedLabel != null)
                        {
                            if (attributeMetadata.Description.UserLocalizedLabel.Label != null)
                            {
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeDescription).Value = attributeMetadata.Description.UserLocalizedLabel.Label;
                            }
                        }
                    }

                    // attribute specicic values
                    switch (attributeMetadata.AttributeType.GetValueOrDefault())
                    {
                        case AttributeTypeCode.Status:
                            {
                                var statusAttributeMetadata = (StatusAttributeMetadata)attributeMetadata;
                                var optionString = String.Empty;
                                foreach (var option in statusAttributeMetadata.OptionSet.Options)
                                {
                                    optionString += option.Value + ":" + option.Label.UserLocalizedLabel.Label;
                                    optionString += ";";
                                }

                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeOptionset).Value = optionString;
                                break;
                            }
                        case AttributeTypeCode.State:
                        {
                                var stateAttributeMetadata = (StateAttributeMetadata)attributeMetadata;
                                var optionString = String.Empty;
                                foreach (var option in stateAttributeMetadata.OptionSet.Options)
                                {
                                    optionString += option.Value + ":" + option.Label.UserLocalizedLabel.Label;
                                    optionString += ";";
                                }

                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeOptionset).Value = optionString;
                                break;
                        }

                        case AttributeTypeCode.Lookup:
                            {
                                var lookupAttributeMetadata = (LookupAttributeMetadata)attributeMetadata;
                                break;
                            }

                        case AttributeTypeCode.String:
                            {
                                var stringAttributeMetadata = (StringAttributeMetadata)attributeMetadata;
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMaxLength).Value = stringAttributeMetadata.MaxLength.ToString();
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeStringFormat).Value = stringAttributeMetadata.FormatName.Value;

                                break;
                            }

                        case AttributeTypeCode.Picklist:
                            {
                                var picklistAttributeMetadata = (PicklistAttributeMetadata)attributeMetadata;
                                var optionString = String.Empty;
                                foreach (var option in picklistAttributeMetadata.OptionSet.Options)
                                {
                                    optionString += option.Value + ":" + option.Label.UserLocalizedLabel.Label;
                                    optionString += ";";
                                }

                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeOptionset).Value = optionString;
                                break;
                            }

                        case AttributeTypeCode.Boolean:
                            {
                                var booleanAttributeMetadata = (BooleanAttributeMetadata)attributeMetadata;
                                var optionString = string.Empty;
                                optionString += booleanAttributeMetadata.OptionSet.TrueOption.Value + ":" + booleanAttributeMetadata.OptionSet.TrueOption.Label.UserLocalizedLabel.Label;
                                optionString += ";";
                                optionString += booleanAttributeMetadata.OptionSet.FalseOption.Value + ":" + booleanAttributeMetadata.OptionSet.FalseOption.Label.UserLocalizedLabel.Label;
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeOptionset).Value = optionString;
                                break;
                            }

                        case AttributeTypeCode.Integer:
                            {
                                var integerAttributeMetadata = (IntegerAttributeMetadata)attributeMetadata;
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMin).Value = integerAttributeMetadata.MinValue?.ToString();
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMax).Value = integerAttributeMetadata.MaxValue?.ToString();
                                break;
                            }

                        case AttributeTypeCode.Money:
                            {
                                var moneyAttributeMetadata = (MoneyAttributeMetadata)attributeMetadata;
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMin).Value = moneyAttributeMetadata.MinValue?.ToString();
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMax).Value = moneyAttributeMetadata.MaxValue?.ToString();
                                break;
                            }

                        case AttributeTypeCode.DateTime:
                            {
                                var dateTimeAttributeMetadata = (DateTimeAttributeMetadata)attributeMetadata;
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeDateBehavior).Value = dateTimeAttributeMetadata.DateTimeBehavior?.Value.ToString();
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeDateFormat).Value = dateTimeAttributeMetadata.Format.ToString();

                                break;
                            }

                        case AttributeTypeCode.Double:
                            {
                                var doubleAttributeMetadata = (DoubleAttributeMetadata)attributeMetadata;
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMin).Value = doubleAttributeMetadata.MinValue?.ToString();
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMax).Value = doubleAttributeMetadata.MaxValue?.ToString();
                                break;
                            }

                        case AttributeTypeCode.Decimal:
                            {
                                var decimalAttributeMetadata = (DecimalAttributeMetadata)attributeMetadata;
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMin).Value = decimalAttributeMetadata.MinValue?.ToString();
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMax).Value = decimalAttributeMetadata.MaxValue?.ToString();
                                break;
                            }

                        case AttributeTypeCode.Memo:
                            {
                                var memoAttributeMetadata = (MemoAttributeMetadata)attributeMetadata;
                                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeMaxLength).Value = memoAttributeMetadata.MaxLength.ToString();
                                break;
                            }
                    }

                    rowAttributeIndex++;
                }
            }
        }

        private void AddAttributeFormInformationToSheet(EntityMetadata entityMetadata, int rowAttributeIndex, AttributeMetadata attributeMetadata)
        {
            if (!this.UseSolutionXml)
            {
                return;
            }

            // FormActivationState
            var locations = string.Empty;
            var active = string.Empty;
            try
            {
                var oAttributeOnForms =
                    this.configurationFile.SelectNodes(
                        $"//Entities/Entity[Name='{entityMetadata.SchemaName}']/FormXml//forms[@type='main']//tabs//tab//sections//section//rows/row//cell//control[@id='{attributeMetadata.SchemaName.ToLower()}']");
                if (oAttributeOnForms == null)
                {
                    return;
                }

                var forms = string.Empty;
                foreach (XmlNode attributeOnForm in oAttributeOnForms)
                {
                    var form = string.Empty;
                    var oCell = attributeOnForm.ParentNode;
                    var labelNode = oCell.SelectSingleNode(".//label[@languagecode='" + this.LanguageCode + "']");

                    if (labelNode == null)
                    {
                        return;
                    }

                    var tab = string.Empty;
                    var sectionDescription = string.Empty;

                    var label = labelNode.Attributes?.GetNamedItem("description").Value;
                    var sectionNode = labelNode.ParentNode?.ParentNode.ParentNode.ParentNode.ParentNode;
                    if (sectionNode != null)
                    {
                        var oTab = sectionNode.ParentNode?.ParentNode?.ParentNode.ParentNode;
                        if (oTab != null)
                        {
                            tab =
                                oTab.SelectSingleNode(".//label[@languagecode='" + this.LanguageCode + "']")?
                                    .Attributes?.GetNamedItem("description")
                                    .Value;
                            if (string.IsNullOrEmpty(tab))
                            {
                                tab = oTab.Attributes?.GetNamedItem("name").Value;
                            }

                            form =
                                oTab.ParentNode?.ParentNode.ParentNode?.SelectSingleNode(
                                    "./LocalizedNames/LocalizedName[@languagecode='" + this.LanguageCode + "']")?
                                    .Attributes?.GetNamedItem("description")
                                    .Value;
                            active =
                                oTab.ParentNode?.ParentNode.ParentNode.SelectSingleNode(".//FormActivationState")
                                    .ChildNodes[0].Value;

                        }

                        sectionDescription =
                            sectionNode.ChildNodes[0].SelectSingleNode(".//label[@languagecode='" + this.LanguageCode +
                                                                       "']")?
                                .Attributes?.GetNamedItem("description")
                                .Value;
                    }

                    //if (active.Equals("1", StringComparison.InvariantCultureIgnoreCase))
                    //{
                    locations = locations + $"{form}-{tab}-{sectionDescription}-{label}\r\n";
                    forms = forms + $"{form}-Active{active}\r\n";
                    //}
                }

                if (!string.IsNullOrEmpty(locations))
                {
                    this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeTab).Value = locations;
                    this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeForm).Value = forms;
                }
            }
            catch (Exception exception)
            {
                this.worksheetAttributes.Cell(rowAttributeIndex, Constants.ColAttributeErrors).Value = exception.Message;
            }
        }
    }
}