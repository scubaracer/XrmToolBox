using System.Diagnostics;
using XrmToolBox.Extensibility;

namespace Daniel.CrmExcel
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.IO;
    using System.Linq;

    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;

    using OfficeOpenXml;

    public class ExcelToCrm : PluginControlBase
    {

        public ExcelToCrm(IOrganizationService service, BackgroundWorker backgroundWorker)
        {
            this.Service = service;
            this.BackgroundWorker = backgroundWorker;
        }

        public string SolutionName
        {
            get;
            private set;
        }

        public int LanguageCode
        {
            get;
            private set;
        }

        public IOrganizationService Service
        {
            get;
        }

        public BackgroundWorker BackgroundWorker
        {
            get;
        }


        public void FixRelations(List<string> Nodes)
        {
            foreach (var node in Nodes)
            {
                Debug.WriteLine($"Entity {node}");
                var entityt = this.RetrieveEntity(node);
                foreach (var oneToManyRelationshipMetadata in entityt.OneToManyRelationships)
                {
                    this.BackgroundWorker.ReportProgress(10, $"Entity {entityt}");
                    if (oneToManyRelationshipMetadata.SchemaName.Contains("wv_relationship_persoon") ||
                        oneToManyRelationshipMetadata.SchemaName.Contains("wv_wv_relationship_wv_rel"))
                    {
                        Debug.WriteLine($"Skipping Entity {node} - relation {oneToManyRelationshipMetadata.SchemaName}");
                    }
                    else
                    {
                        if (oneToManyRelationshipMetadata.CascadeConfiguration.Delete == CascadeType.RemoveLink ||
                            (oneToManyRelationshipMetadata.SchemaName.Contains("wv_family_wv_familymember_familyid") && oneToManyRelationshipMetadata.CascadeConfiguration.Delete == CascadeType.Cascade))
                        {
                            oneToManyRelationshipMetadata.CascadeConfiguration.Delete = CascadeType.Restrict;
                            this.BackgroundWorker.ReportProgress(10, $"Relation {oneToManyRelationshipMetadata.SchemaName}");
                            try
                            {
                                var updateRelationShip = new UpdateRelationshipRequest
                                {
                                    Relationship = oneToManyRelationshipMetadata,
                                };
                                Debug.WriteLine($"Entity {node} - relation {oneToManyRelationshipMetadata.SchemaName}");
                                this.Service.Execute(updateRelationShip);
                            }
                            catch (Exception exception)
                            {
                                Debug.WriteLine($"Entity {node} - relation {oneToManyRelationshipMetadata.SchemaName} error {exception.Message}");
                                var errror = exception.Message;
                            }
                        }

                    }
                }

            }
            Debug.WriteLine($"Ready");
        }

        public void Start(string excelFile, string solutionName, int languageCode)
        {
            this.SolutionName = solutionName;
            this.LanguageCode = languageCode;
            this.BackgroundWorker.ReportProgress(10, "Reading Excel file");
            // 1 open excel
            // 2 read attributes
            // 3    retrieve entity and check if attribute already exists
            // 4    if attribute is new create it!
            var crmExcel = new FileInfo(excelFile);
            // else create entity
            using (var excelPackage = new ExcelPackage(crmExcel))
            {

                // Get handle to the existing worksheet
                var worksheetEntities = excelPackage.Workbook.Worksheets["Entities"];
                var worksheetAttributes = excelPackage.Workbook.Worksheets["Attributes"];
                var worksheetRelationShipsManyToOne = excelPackage.Workbook.Worksheets["RelationShips Many to One"];
                var worksheetRelationShipsManyToMany = excelPackage.Workbook.Worksheets["RelationShips Many to Many"];
                var entityMetadataCollection = new Dictionary<string, EntityMetadata>();
                if (worksheetEntities != null)
                {
                    var currentRow = 2;
                    while (true)
                    {
                        if (worksheetEntities.Cell(currentRow, Constants.ColSchemaName).Value == string.Empty)
                        {
                            break;
                        }

                        var entityLogicalName = worksheetEntities.Cell(currentRow, Constants.ColSchemaName).Value.ToLower();
                        this.BackgroundWorker.ReportProgress(10, $"Processing entity '{entityLogicalName}'");
                        var entity = this.RetrieveEntity(entityLogicalName);
                        if (entity == null)
                        {
                            this.CreateEntityFromSheet(entityLogicalName, worksheetEntities, currentRow, worksheetAttributes);
                        }

                        currentRow++;
                    }
                }

                this.CreateAttributes(worksheetAttributes, entityMetadataCollection);
                this.CreateManyToOneRelations(entityMetadataCollection, worksheetRelationShipsManyToOne);
                this.CreateManyToManyRelationships(entityMetadataCollection, worksheetRelationShipsManyToMany);
                excelPackage.Save(); // this to save the errors that have occurred.
                this.BackgroundWorker.ReportProgress(10, "Ready");
            }
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
                LogError($"Error retrieving entityt metadata for entity '{logicalName}'",exception);
            }

            return null;
        }

        private void CreateEntityFromSheet(string entityLogicalName, ExcelWorksheet worksheetEntities, int currentRow, ExcelWorksheet worksheetAttributes)
        {
            this.BackgroundWorker.ReportProgress(10, $"Processing entity '{entityLogicalName}' - Creating");
            var createEntityRequest = new CreateEntityRequest
                                          {
                                              Entity = new EntityMetadata
                                                           {
                                                               SchemaName = entityLogicalName,
                                                               DisplayName = new Label(worksheetEntities.Cell(currentRow, Constants.ColEntityDisplayName).Value, this.LanguageCode),
                                                               DisplayCollectionName = new Label(worksheetEntities.Cell(currentRow, Constants.ColDisplayCollectionName).Value, this.LanguageCode),
                                                               Description = new Label(worksheetEntities.Cell(currentRow, Constants.ColDescription).Value, this.LanguageCode),
                                                               OwnershipType = EnumHelper.ParseByDescription<OwnershipTypes>(worksheetEntities.Cell(currentRow, Constants.ColOwnerShipType).Value),
                                                               IsActivity = false
                                                           },

                                              // Define the primary attribute for the entity
                                              PrimaryAttribute = new StringAttributeMetadata
                                                                     {
                                                                         SchemaName = worksheetEntities.Cell(currentRow, Constants.ColPrimaryAttributeSchemaName).Value,
                                                                         RequiredLevel = new AttributeRequiredLevelManagedProperty(EnumHelper.ParseByDescription<AttributeRequiredLevel>(worksheetAttributes.Cell(currentRow, Constants.ColPrimaryAttributeRequired).Value)),
                                                                         MaxLength = Convert.ToInt32(worksheetEntities.Cell(currentRow, Constants.ColPrimaryAttributeLength).Value),
                                                                         Format = StringFormat.Text,
                                                                         DisplayName = new Label(worksheetEntities.Cell(currentRow, Constants.ColPrimaryAttributeDisplayname).Value, this.LanguageCode),
                                                                         Description = new Label(worksheetEntities.Cell(currentRow, Constants.ColPrimaryAttributeDescription).Value, this.LanguageCode)
                                                                     },
                                              SolutionUniqueName = this.SolutionName
                                          };
            try
            {
                LogInfo($"Create entity {entityLogicalName}");
               this.Service.Execute(createEntityRequest);
            }
            catch (Exception exception)
            {
                worksheetAttributes.Cell(currentRow, Constants.ColEntityError).Value = exception.Message;
            }
        }

        private void CreateManyToManyRelationships(Dictionary<string, EntityMetadata> entityMetadataCollection, ExcelWorksheet worksheetRelationShipsManyToMany)
        {
            if (worksheetRelationShipsManyToMany != null)
            {
                var currentRow = 2;
                while (true)
                {
                    var entityLogicalName = worksheetRelationShipsManyToMany.Cell(currentRow, Constants.ColSchemaName).Value.ToLower();
                    if (entityLogicalName == string.Empty)
                    {
                        break;
                    }

                    var entityMetaData = this.GetEntityMetaData(entityMetadataCollection, entityLogicalName);

                    var schemaName = worksheetRelationShipsManyToMany.Cell(currentRow, Constants.ColRelationShipSchemaName).Value;
                    var relationship = entityMetaData.ManyToManyRelationships.FirstOrDefault(a => a.SchemaName.ToLower() == schemaName.ToLower());
                    if (relationship == null)
                    {
                        var manyToManyRelationshipMetadata = new ManyToManyRelationshipMetadata
                                                                 {
                                                                     SchemaName = worksheetRelationShipsManyToMany.Cell(currentRow, Constants.ColRelationShipSchemaName).Value.ToLower(),
                                                                     Entity1LogicalName = worksheetRelationShipsManyToMany.Cell(currentRow, Constants.ColRelationShipMmEntity1LogicalName).Value.ToLower(),
                                                                     Entity2LogicalName = worksheetRelationShipsManyToMany.Cell(currentRow, Constants.ColRelationShipMmEntity2LogicalName).Value.ToLower(),
                                                                     Entity1AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                                                                                                              {
                                                                                                                  Behavior = AssociatedMenuBehavior.UseLabel,
                                                                                                                  Group = AssociatedMenuGroup.Details,
                                                                                                                  Label = new Label(worksheetRelationShipsManyToMany.Cell(currentRow, Constants.ColRelationShipMmEntity1AssociatedMenuLabel).Value, this.LanguageCode),
                                                                                                                  Order = 10000
                                                                                                              },
                                                                     Entity2AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                                                                                                              {
                                                                                                                  Behavior = AssociatedMenuBehavior.UseLabel,
                                                                                                                  Group = AssociatedMenuGroup.Details,
                                                                                                                  Label = new Label(worksheetRelationShipsManyToMany.Cell(currentRow, Constants.ColRelationShipMmEntity2AssociatedMenuLabel).Value, this.LanguageCode),
                                                                                                                  Order = 10000
                                                                                                              }
                                                                 };

                        var createManyToManyRequest = new CreateManyToManyRequest
                                                          {
                                                              ManyToManyRelationship = manyToManyRelationshipMetadata,
                                                              SolutionUniqueName = this.SolutionName,
                                                              IntersectEntitySchemaName = worksheetRelationShipsManyToMany.Cell(currentRow, Constants.ColRelationShipMmIntersectEntitySchemaName).Value.ToLower()
                                                          };
                        try
                        {
                            LogInfo($"Create relationship M:M {manyToManyRelationshipMetadata.SchemaName} on entity {entityLogicalName}");
                            this.Service.Execute(createManyToManyRequest);
                        }
                        catch (Exception exception)
                        {
                            worksheetRelationShipsManyToMany.Cell(currentRow, Constants.ColRelationShipErrors).Value = exception.Message;
                        }
                    }

                    currentRow++;
                }
            }
        }

        private EntityMetadata GetEntityMetaData(Dictionary<string, EntityMetadata> entityMetadataCollection, string entityLogicalName)
        {
            EntityMetadata entityMetaData;
            if (entityMetadataCollection.ContainsKey(entityLogicalName))
            {
                entityMetaData = entityMetadataCollection[entityLogicalName];
            }
            else
            {
                entityMetaData = this.RetrieveEntity(entityLogicalName);
                entityMetadataCollection.Add(entityLogicalName, entityMetaData);
            }

            return entityMetaData;
        }

        private void CreateAttributes(ExcelWorksheet worksheetAttributes, Dictionary<string, EntityMetadata> entityMetadataCollection)
        {
            if (worksheetAttributes != null)
            {
                var currentRow = 2;
                while (true)
                {
                    try
                    {
                        var entityLogicalName = worksheetAttributes.Cell(currentRow, Constants.ColSchemaName).Value.ToLower();
                        if (entityLogicalName == string.Empty)
                        {
                            break;
                        }

                        var entityMetaData = this.GetEntityMetaData(entityMetadataCollection, entityLogicalName);

                        if (entityMetaData == null)
                        {
                            return;
                        }

                        var schemaName = worksheetAttributes.Cell(currentRow, Constants.ColAttributeSchemaName).Value;
                        var attribute = entityMetaData.Attributes.FirstOrDefault(a => a.SchemaName == schemaName);
                        if (attribute != null)
                        {
                            // Existing attribute, do nothing 
                            // Only change the Displayname
                            var maxLengthDiffers = false;
                            switch (attribute.AttributeType)
                            {
                                case AttributeTypeCode.String:
                                    {
                                        var stringAttribute = (StringAttributeMetadata)attribute;
                                        if (stringAttribute.MaxLength.Value != Convert.ToInt32(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMaxLength).Value))
                                        {
                                            maxLengthDiffers = true;
                                        }

                                        break;
                                    }
                            }

                            var displayname = worksheetAttributes.Cell(currentRow, Constants.ColAttributeDisplayName).Value;
                            if (((attribute.DisplayName.UserLocalizedLabel != null && attribute.DisplayName.UserLocalizedLabel.Label != displayname) || attribute.RequiredLevel.Value != EnumHelper.ParseByDescription<AttributeRequiredLevel>(worksheetAttributes.Cell(currentRow, Constants.ColAttributeRequired).Value) || attribute.IsAuditEnabled.Value != Convert.ToBoolean(worksheetAttributes.Cell(currentRow, Constants.ColAttributeIsAuditEnabled).Value) || attribute.IsValidForAdvancedFind.Value != Convert.ToBoolean(worksheetAttributes.Cell(currentRow, Constants.ColAttributeIsValidForAdvancedFind).Value) || maxLengthDiffers) && attribute.IsPrimaryId.Value != true)
                            {
                                if (attribute.DisplayName.UserLocalizedLabel != null && attribute.DisplayName.UserLocalizedLabel?.Label != displayname)
                                {
                                    attribute.DisplayName = new Label(displayname, this.LanguageCode);
                                }

                                if (attribute.RequiredLevel.Value != EnumHelper.ParseByDescription<AttributeRequiredLevel>(worksheetAttributes.Cell(currentRow, Constants.ColAttributeRequired).Value))
                                {
                                    attribute.RequiredLevel = new AttributeRequiredLevelManagedProperty(EnumHelper.ParseByDescription<AttributeRequiredLevel>(worksheetAttributes.Cell(currentRow, Constants.ColAttributeRequired).Value));
                                }

                                if (attribute.IsAuditEnabled.Value != Convert.ToBoolean(worksheetAttributes.Cell(currentRow, Constants.ColAttributeIsAuditEnabled).Value))
                                {
                                    attribute.IsAuditEnabled = new BooleanManagedProperty(Convert.ToBoolean(worksheetAttributes.Cell(currentRow, Constants.ColAttributeIsAuditEnabled).Value));
                                }

                                if (attribute.IsValidForAdvancedFind.Value != Convert.ToBoolean(worksheetAttributes.Cell(currentRow, Constants.ColAttributeIsValidForAdvancedFind).Value))
                                {
                                    attribute.IsValidForAdvancedFind = new BooleanManagedProperty(Convert.ToBoolean(worksheetAttributes.Cell(currentRow, Constants.ColAttributeIsValidForAdvancedFind).Value));
                                }

                                if (maxLengthDiffers)
                                {
                                    switch (attribute.AttributeType)
                                    {
                                        case AttributeTypeCode.String:
                                            {
                                                var stringAttribute = (StringAttributeMetadata)attribute;
                                                stringAttribute.MaxLength = Convert.ToInt32(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMaxLength).Value);
                                                break;
                                            }
                                    }
                                }

                                var updateAttributeRequest = new UpdateAttributeRequest
                                                                 {
                                                                     Attribute = attribute,
                                                                     EntityName = entityLogicalName,
                                                                     SolutionUniqueName = this.SolutionName
                                                                 };
                                try
                                {
                                    this.BackgroundWorker.ReportProgress(10, $"Processing entity '{entityLogicalName}' - Updateing attribute '{schemaName}'");
                                    LogInfo($"Update attribute {schemaName} on entity {entityLogicalName}");
                                    this.Service.Execute(updateAttributeRequest);
                                }
                                catch (Exception exception)
                                {
                                    worksheetAttributes.Cell(currentRow, Constants.ColAttributeErrors).Value = exception.Message;
                                }
                            }
                        }
                        else
                        {
                            // create a new one!
                            // Get type
                            var attributeTypeCode = EnumHelper.ParseByDescription<AttributeTypeCode>(worksheetAttributes.Cell(currentRow, Constants.ColAttributeType).Value);
                            var displayName = worksheetAttributes.Cell(currentRow, Constants.ColAttributeDisplayName).Value;
                            var description = worksheetAttributes.Cell(currentRow, Constants.ColAttributeDescription).Value;
                            AttributeMetadata newAttribute = null;
                            switch (attributeTypeCode)
                            {
                                case AttributeTypeCode.Boolean:
                                    {
                                        // Create a boolean attribute
                                        var options = worksheetAttributes.Cell(currentRow, Constants.ColAttributeOptionset).Value.Split(";".ToCharArray());
                                        var optionTrue = options[0].Split(":".ToCharArray());
                                        var optionFalse = options[1].Split(":".ToCharArray());
                                        var boolAttribute = new BooleanAttributeMetadata
                                                                {
                                                                    OptionSet = new BooleanOptionSetMetadata(new OptionMetadata(new Label(optionTrue[1], this.LanguageCode), Convert.ToInt32(optionTrue[0])), new OptionMetadata(new Label(optionFalse[1], this.LanguageCode), Convert.ToInt32(optionFalse[0])))
                                                                };
                                        newAttribute = boolAttribute;
                                        break;
                                    }

                                case AttributeTypeCode.DateTime:
                                    {
                                        // Create a date time attribute
                                        var dtAttribute = new DateTimeAttributeMetadata
                                                              {
                                                                  Format = DateTimeFormat.DateOnly,
                                                                  ImeMode = Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled
                                        };
                                        newAttribute = dtAttribute;
                                        break;
                                    }

                                case AttributeTypeCode.Decimal:
                                    {
                                        // Create a decimal attribute
                                        var decimalAttribute = new DecimalAttributeMetadata
                                                                   {
                                                                       MaxValue = Convert.ToInt32(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMax).Value),
                                                                       MinValue = Convert.ToInt32(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMin).Value),
                                                                       Precision = 1
                                                                   };
                                        newAttribute = decimalAttribute;
                                        break;
                                    }

                                case AttributeTypeCode.Integer:
                                    {
                                        // Create a integer attribute
                                        var integerAttribute = new IntegerAttributeMetadata
                                                                   {
                                                                       Format = IntegerFormat.None
                                                                   };
                                        if (worksheetAttributes.Cell(currentRow, Constants.ColAttributeMin).Value != string.Empty)
                                        {
                                            integerAttribute.MinValue = Convert.ToInt32(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMin).Value);
                                        }

                                        if (worksheetAttributes.Cell(currentRow, Constants.ColAttributeMax).Value != string.Empty)
                                        {
                                            integerAttribute.MaxValue = Convert.ToInt32(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMax).Value);
                                        }

                                        newAttribute = integerAttribute;
                                        break;
                                    }

                                case AttributeTypeCode.Memo:
                                    {
                                        // Create a memo attribute
                                        var memoAttribute = new MemoAttributeMetadata
                                                                {
                                                                    Format = StringFormat.TextArea,
                                                                    ImeMode = Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled,
                                                                    MaxLength = Convert.ToInt32(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMaxLength).Value)
                                                                };
                                        newAttribute = memoAttribute;
                                        break;
                                    }

                                case AttributeTypeCode.Money:
                                    {
                                        // Create a money attribute
                                        var moneyAttribute = new MoneyAttributeMetadata
                                                                 {
                                                                     // Set extended properties
                                                                     MaxValue = Convert.ToDouble(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMax).Value),
                                                                     MinValue = Convert.ToDouble(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMin).Value),
                                                                     Precision = 1,
                                                                     PrecisionSource = 1,
                                                                     ImeMode = Microsoft.Xrm.Sdk.Metadata.ImeMode.Disabled
                                        };
                                        newAttribute = moneyAttribute;
                                        break;
                                    }

                                case AttributeTypeCode.Picklist:
                                    {
                                        var options = worksheetAttributes.Cell(currentRow, Constants.ColAttributeOptionset).Value.Split(";".ToCharArray());
                                        var optionSet = new OptionSetMetadata
                                                            {
                                                                IsGlobal = false,
                                                                OptionSetType = OptionSetType.Picklist
                                                            };
                                        foreach (var option in options)
                                        {
                                            if (option != string.Empty)
                                            {
                                                var optionMetaData = new OptionMetadata();
                                                var optionData = option.Split(":".ToCharArray());
                                                optionMetaData.Label = new Label(optionData[1], this.LanguageCode);
                                                optionMetaData.Value = Convert.ToInt32(optionData[0]);
                                                optionSet.Options.Add(optionMetaData);
                                            }
                                        }

                                        var pickListAttribute = new PicklistAttributeMetadata
                                                                    {
                                                                        OptionSet = optionSet
                                                                    };
                                        newAttribute = pickListAttribute;
                                        break;
                                    }

                                case AttributeTypeCode.String:
                                    {
                                        // Create a string attribute
                                        var stringAttribute = new StringAttributeMetadata
                                                                  {
                                                                        MaxLength = Convert.ToInt32(worksheetAttributes.Cell(currentRow, Constants.ColAttributeMaxLength).Value),
                                                                        Format = EnumHelper.ParseByDescription<StringFormat>(worksheetAttributes.Cell(currentRow, Constants.ColAttributeStringFormat).Value)
                                                                };
                                        newAttribute = stringAttribute;
                                        break;
                                    }
                            }
                            if (newAttribute != null)
                            {
                                // Create the request.
                                newAttribute.SchemaName = schemaName;
                                newAttribute.DisplayName = new Label(displayName, this.LanguageCode);
                                newAttribute.Description = new Label(description, this.LanguageCode);
                                newAttribute.RequiredLevel = new AttributeRequiredLevelManagedProperty(EnumHelper.ParseByDescription<AttributeRequiredLevel>(worksheetAttributes.Cell(currentRow, Constants.ColAttributeRequired).Value));
                                newAttribute.IsAuditEnabled = new BooleanManagedProperty(Convert.ToBoolean(worksheetAttributes.Cell(currentRow, Constants.ColAttributeIsAuditEnabled).Value));
                                newAttribute.IsValidForAdvancedFind = new BooleanManagedProperty(Convert.ToBoolean(worksheetAttributes.Cell(currentRow, Constants.ColAttributeIsValidForAdvancedFind).Value));

                                var createAttributeRequest = new CreateAttributeRequest
                                                                 {
                                                                     EntityName = entityLogicalName,
                                                                     Attribute = newAttribute,
                                                                     SolutionUniqueName = this.SolutionName
                                                                 };

                                // Execute the request.
                                try
                                {
                                    this.BackgroundWorker.ReportProgress(10, $"Processing entity '{entityLogicalName}' - Creating attribute '{schemaName}'");
                                    LogInfo($"Create attribute {schemaName} on entity {entityLogicalName}");
                                    this.Service.Execute(createAttributeRequest);
                                }
                                catch (Exception exception)
                                {
                                    worksheetAttributes.Cell(currentRow, Constants.ColAttributeErrors).Value = exception.Message;
                                }
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        worksheetAttributes.Cell(currentRow, Constants.ColAttributeErrors).Value = exception.Message;
                    }

                    currentRow++;
                }
            }
        }

        private void CreateManyToOneRelations(Dictionary<string, EntityMetadata> entityMetadataCollection, ExcelWorksheet worksheetRelationShipsManyToOne)
        {
            if (worksheetRelationShipsManyToOne != null)
            {
                var currentRow = 2;
                while (true)
                {
                    try
                    {
                        var entityLogicalName = worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColSchemaName).Value.ToLower();
                        if (entityLogicalName == string.Empty)
                        {
                            break;
                        }

                        var entityMetaData = this.GetEntityMetaData(entityMetadataCollection, entityLogicalName);

                        var schemaName = worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipSchemaName).Value;
                        var addRelationShip = true;
                        var res = entityMetaData.ManyToOneRelationships.FirstOrDefault(a => a.SchemaName.ToLower() == schemaName.ToLower());

                        if (res != null)
                        {
                            addRelationShip = false;
                        }

                        res = entityMetaData.OneToManyRelationships.FirstOrDefault(a => a.SchemaName.ToLower() == schemaName.ToLower());
                        if (res != null)
                        {
                            addRelationShip = false;
                        }

                        if (addRelationShip)
                        {
                            var oneToManyRelationshipMetadata = new OneToManyRelationshipMetadata();
                            oneToManyRelationshipMetadata.SchemaName = worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipSchemaName).Value.ToLower();
                            oneToManyRelationshipMetadata.ReferencedEntity = worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipReferencedEntity).Value.ToLower();
                            oneToManyRelationshipMetadata.ReferencingEntity = worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipReferencingEntity).Value.ToLower();

                            oneToManyRelationshipMetadata.AssociatedMenuConfiguration = new AssociatedMenuConfiguration
                                                                                            {
                                 //Behavior = AssociatedMenuBehavior.UseLabel,
                                Behavior = AssociatedMenuBehavior.UseCollectionName,
                                Group = AssociatedMenuGroup.Details,
                                   //                                                             Label = new Label(worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipLabel).Value, LanguageCode),
                                                                                                Order = 10000
                                                                                            };
                            // oneToManyRelationshipMetadata.CascadeConfiguration = new CascadeConfiguration  {Assign = CascadeType.Cascade,Delete = CascadeType.Cascade,Merge = CascadeType.Cascade,Reparent = CascadeType.Cascade,Share = CascadeType.Cascade,Unshare = CascadeType.Cascade};
                            var lookup = new LookupAttributeMetadata
                                             {
                                                 SchemaName = worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipLookUpSchemaName).Value.ToLower(),
                                                 DisplayName = new Label(worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipLookUpDisplayName).Value, this.LanguageCode),
                                                 RequiredLevel = new AttributeRequiredLevelManagedProperty(EnumHelper.ParseByDescription<AttributeRequiredLevel>(worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipLookUpRequiredLevel).Value)),
                                                 Description = new Label(worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipLookUpDescription).Value, this.LanguageCode)
                                             };

                            var createOneToManyRequest = new CreateOneToManyRequest
                                                             {
                                                                 OneToManyRelationship = oneToManyRelationshipMetadata,
                                                                 SolutionUniqueName = this.SolutionName,
                                                                 Lookup = lookup
                                                             };
                            try
                            {
                                LogInfo($"Create relationship {oneToManyRelationshipMetadata.SchemaName} on entity {entityLogicalName}");
                                this.Service.Execute(createOneToManyRequest);
                            }
                            catch (Exception exception)
                            {
                                worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipErrors).Value = exception.Message;
                            }
                        }
                    }
                    catch (Exception exception)
                    {
                        worksheetRelationShipsManyToOne.Cell(currentRow, Constants.ColRelationShipErrors).Value = exception.Message;
                    }

                    currentRow++;
                }
            }
        }
    }
}