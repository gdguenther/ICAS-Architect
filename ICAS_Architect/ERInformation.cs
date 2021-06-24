using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using Microsoft.Msagl.Drawing;
using System.Data.SqlClient;
using Dapper;

namespace ICAS_Architect
{
    //
    // ERInformation captures Entity information (EREntityAttribute) and Relationship information (ERRelation) for rendering and serialization purposes.
    // ERInformation.ConvertToGraph() creates a Graph object for MSAGL for selected entities for visualization.
    //

    [JsonObject]
    internal class EREntity
    {
        // don't bother writing these to a file, they can be generated themselves
        internal string organizationUrl;
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal string entityDefinitionsPath;
        internal string manyToManyRelationshipsPath;
        internal readonly string AttributesUrl;
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal readonly string EntityListUrl;

        [JsonProperty]        internal string metadataId;
        [JsonProperty]        internal string entityLogicalName;
        [JsonProperty] internal string entitySetName;
        [JsonProperty] internal string entitySchema;
        [JsonProperty]        internal string descriptiontext;
        [JsonProperty]        internal int objectTypeCode;
        [JsonProperty]        internal string primaryIdAttribute;
        [JsonProperty]        internal string primaryNameAttribute;
        [JsonProperty]        internal string displayNamesText;
        [JsonProperty] internal bool isIntersect;
        [JsonProperty] internal string databaseName;
        [JsonProperty] internal string Table_Type;

        internal EREntity()
        {
        }

        internal EREntity(string organizationUrl, string APIURL, string metadataId, string entityLogicalName, string entitySetName, string descriptiontext, int objectTypeCode, string primaryIdAttribute, string primaryNameAttribute, string displayNamesText, bool isIntersect, string Table_Type, string databaseName = "", string schema = "")
        {
            this.organizationUrl = organizationUrl;
            this.entityDefinitionsPath = $"{APIURL}/EntityDefinitions({metadataId})";
            this.manyToManyRelationshipsPath = $"{entityDefinitionsPath}/ManyToManyRelationships?$select=Entity1LogicalName,Entity2LogicalName,Entity1IntersectAttribute,Entity2IntersectAttribute,IntersectEntityName,SchemaName";
            this.AttributesUrl = $"{this.organizationUrl}{entityDefinitionsPath}/Attributes?$select=MetadataId,SchemaName,AttributeTypeName,DisplayName";
            this.EntityListUrl = $"{this.organizationUrl}/main.aspx?etn={entityLogicalName}&pagetype=entitylist";
            this.metadataId = metadataId;
            this.entityLogicalName = entityLogicalName;
            this.entitySetName = entitySetName;
            this.descriptiontext = descriptiontext;
            this.objectTypeCode = objectTypeCode;
            this.primaryIdAttribute = primaryIdAttribute;
            this.primaryNameAttribute = primaryNameAttribute;
            this.displayNamesText = displayNamesText;
            this.isIntersect = isIntersect;
            this.databaseName = databaseName;
            this.entitySchema = schema;
            this.Table_Type = Table_Type;
        }
    }

    [JsonObject]
    internal class EREntityAttribute
    {
        [JsonProperty]        internal string EntityName;
        [JsonProperty]        internal string AttributeName;
        [JsonProperty]        internal string DataType;
        [JsonProperty]        internal string Description;
        [JsonProperty]        internal int ColumnNumber;
        [JsonProperty]        internal string EntityLogicalName;
        [JsonProperty]        internal bool IsNullable;
        [JsonProperty]        internal bool IsPrimaryID;
        [JsonProperty]        internal string AttributeType;
        [JsonProperty]        internal string MetadataId;
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal string AttributeOne;    // Intersect entity. If name is set, EntityOne and EntityMany are "Many-To-Many".
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal string Default;

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal int? CHARACTER_MAXIMUM_LENGTH;
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal int? CHARACTER_OCTET_LENGTH;
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal int? NUMERIC_PRECISION;
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal int? NUMERIC_PRECISION_RADIX;
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal int? NUMERIC_SCALE;
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal int? DATETIME_PRECISION;

        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal string Character_Set_Catalog;    // Intersect entity. If name is set, EntityOne and EntityMany are "Many-To-Many".
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal string Character_Set_Schema;    // Intersect entity. If name is set, EntityOne and EntityMany are "Many-To-Many".
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal string Collation_Catalog;    // Intersect entity. If name is set, EntityOne and EntityMany are "Many-To-Many".
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal string Collation_Schema;    // Intersect entity. If name is set, EntityOne and EntityMany are "Many-To-Many".
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal string Collation_Name;    // Intersect entity. If name is set, EntityOne and EntityMany are "Many-To-Many".

        internal EREntityAttribute()
        {
        }

        internal EREntityAttribute(string entityName, string attributeName, string dataType, string description, int columnNumber, string entityLogicalName, bool isPrimaryID, string attributeType, string metadataId)
//        internal EREntityAttribute(string entityName, string attributeName, string dataType, string description)
        {
            EntityName = entityName;
            AttributeName = attributeName;
            DataType = dataType;
            Description = description;
            ColumnNumber = columnNumber;
            EntityLogicalName = entityLogicalName;
            IsPrimaryID = isPrimaryID;
            AttributeType = attributeType;
            MetadataId = metadataId;
        }

        internal EREntityAttribute(string entityName, string attributeName, string dataType, string description, int columnNumber, string entityLogicalName, bool isPrimaryID, string attributeType, string metadataId, string attributeOne, int? CHARACTER_MAXIMUM_LENGTH = null, int? CHARACTER_OCTET_LENGTH = null, int? NUMERIC_PRECISION = null, int? NUMERIC_PRECISION_RADIX = null, int? NUMERIC_SCALE = null, int? DATETIME_PRECISION = null)
        //        internal EREntityAttribute(string entityName, string attributeName, string dataType, string description)
        {
            EntityName = entityName;
            AttributeName = attributeName;
            DataType = dataType;
            Description = description;
            ColumnNumber = columnNumber;
            EntityLogicalName = entityLogicalName;
            IsPrimaryID = isPrimaryID;
            AttributeType = attributeType;
            MetadataId = metadataId;
            AttributeOne = attributeOne;
            this.CHARACTER_MAXIMUM_LENGTH = CHARACTER_MAXIMUM_LENGTH;
            this.CHARACTER_OCTET_LENGTH = CHARACTER_OCTET_LENGTH;
            this.NUMERIC_PRECISION = NUMERIC_PRECISION;
            this.NUMERIC_PRECISION_RADIX = NUMERIC_PRECISION_RADIX;
            this.NUMERIC_SCALE = NUMERIC_SCALE;
            this.DATETIME_PRECISION = DATETIME_PRECISION;

            this.IsNullable = false;
            this.Collation_Catalog = null;
            this.Collation_Name = null;
            this.Collation_Schema = null;
            this.Character_Set_Catalog = null;
            this.Character_Set_Schema = null;
            this.Default = null;
        }
    }

    [JsonObject]
    internal class ERRelation
    {
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)] internal string RelationName;    
        [JsonProperty]        internal string EntityOne;
        [JsonProperty]        internal string EntityMany;
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]        internal string AttributeOne; 
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]        internal string AttributeMany;        // Lookup field name from the many entity side.
        [JsonProperty(NullValueHandling = NullValueHandling.Ignore)]        internal string IntersectEntity;    // Intersect entity. If name is set, EntityOne and EntityMany are "Many-To-Many".


        internal ERRelation()
        {
        }

        internal ERRelation(string entityOne, string entityMany, string attributeMany)
        {
            EntityOne = entityOne;
            EntityMany = entityMany;
            AttributeMany = attributeMany;
        }

        internal ERRelation(string entityMany1, string entityMany2, string intersectEntity, bool isIntersect)
        {
            if (!isIntersect) throw new InvalidOperationException("This constructor is only for intersect relationship/many-to-many.");
            EntityOne = entityMany1;
            EntityMany = entityMany2;
            IntersectEntity = intersectEntity;
        }
        internal ERRelation(string entityMany1, string entityMany2, string intersectEntity, string relationName, string attributeOne)
        {
            EntityOne = entityMany1;
            EntityMany = entityMany2;
            IntersectEntity = intersectEntity;
            AttributeOne = attributeOne;
            RelationName = relationName;
        }

        internal bool IsManyToMany()
        {
            return (IntersectEntity != null);
        }
    }

    [JsonObject]
    internal class ERInformation
    {
        // list of common entities that are usually 'too much' to render in the graph
        private const string SKIP_ENTITIES_FOR_GRAPH = @",businessunit,organization,systemuser,team,transactioncurrency,workflow,webresource,webwizard,uom,transformationparametermapping,transformationmapping,traceregarding,tracelog,topicmodelconfiguration,topicmodel,timezonedefinition,template,teamtemplate,syncattributemappingprofile,syncattributemapping,subscription,solution,dependencynode,solution,solutioncomponent,sdkmessage,sdkmessagefilter,sdkmessagepair,sdkmessageprocessingstep,sdkmessageprocessingstepsecureconfig,sdkmessagerequest,sdkmessageresponse,role,roletemplate,rollupproperties,routingrule,ribboncustomization,ribbondiff,report,relationshiprole,recurringappointmentmaster,recurrencerule,recommendationmodel,recommendationmodelmapping,recommendationmodelversion,ratingmodel,ratingvalue,principalsyncattributemap,privilege,pluginassembly,plugintype,Empty,asyncoperation,processtrigger,systemform,attributemap,entitymap,knowledgesearchmodel,azureserviceconnection,advancedsimilarityrule,teamprofiles,fieldsecurityprofile,systemuserprofiles,fieldpermission,mobileofflineprofileitemassociation,mobileofflineprofileitem,mobileofflineprofile,savedquery,publisher,publisheraddress,appmodule,appmoduleroles,appmodulecomponent,importlog,importdata,importfile,import,ownermapping,importentitymapping,importmap,columnmapping,lookupmapping,picklistmapping,";

        [JsonProperty]        internal string OrganizationUrl;
        [JsonProperty]        internal DateTime CreatedOn;
        [JsonProperty]        internal string DataSourcePath;
        [JsonProperty]        internal List<EREntity> EREntities;
        [JsonProperty]        internal List<ERRelation> ERRelations;
        [JsonProperty]        internal List<EREntityAttribute> EREntitieAttributes;

        public ERInformation()
        {
            EREntities = new List<EREntity>();
            EREntitieAttributes = new List<EREntityAttribute>();
            ERRelations = new List<ERRelation>();
        }

        internal void AddMetadata(ERInformation erInformationToBeAdded)
        {
            EREntities.AddRange(erInformationToBeAdded.EREntities);
            EREntitieAttributes.AddRange(erInformationToBeAdded.EREntitieAttributes);
            ERRelations.AddRange(erInformationToBeAdded.ERRelations);
        }

        internal IEnumerable<string> ListEntities()
        {
            return ERRelations.Select(rel => rel.EntityOne).Union(ERRelations.Select(rel => rel.EntityMany)).Distinct().OrderBy(val => val);
        }

        internal Graph ConvertToGraph(List<string> targetEntities, int depth)
        {
            // Populate MSAGL Graph object for the selected entities.
            // Find related entities as well if depth is set larger than 0 (i.e. 1 means related entities. 2+ depth is too large to use practically)
            ExpandTargetEntities(targetEntities, depth);

            Graph graph = new Graph("graph");
            foreach (ERRelation relation in ERRelations.Distinct().OrderBy(obj => obj.EntityOne))
            {
                if (!targetEntities.Contains(relation.EntityMany) || !targetEntities.Contains(relation.EntityOne)) continue;
                AddNewRelation(graph, relation);
            }
            return graph;
        }

        private static void AddNewRelation(Graph graph, ERRelation relation)
        {
            string attrId = (relation.IsManyToMany() ? $"{relation.EntityMany} n-{relation.IntersectEntity}-n {relation.EntityOne}" : $"{relation.EntityMany}.{relation.AttributeMany} n-1 {relation.EntityOne}");
            System.Diagnostics.Debug.WriteLine(attrId);
            if (graph.EdgeById(attrId) != null) return;
            Edge edge = graph.AddEdge(relation.EntityMany, relation.EntityOne);
            edge.Attr.ArrowheadAtSource = ArrowStyle.Diamond;
            edge.Attr.ArrowheadAtTarget = (relation.IsManyToMany() ? ArrowStyle.Diamond : ArrowStyle.None);
            edge.Attr.Id = attrId;
        }

        internal string FindParentTableID(string tableName)
        {
            foreach(EREntity eREntity in EREntities)
            {
                if(eREntity.entityLogicalName == tableName)
                {
                    return (eREntity.metadataId);
                }
            }
            return null;
        }

        internal List<string> FindPrimaryKeys(string targetEntity)
        {
            var returnList = EREntitieAttributes.FindAll(x => x.EntityName == targetEntity);
            return null;
        }

        internal List<string> FindRelatedEntities(string targetEntity)
        {
            List<string> relatedEntities = new List<string>();
            relatedEntities.Add(targetEntity);
            foreach (ERRelation relation in ERRelations.Distinct())
            {
                if (relation.EntityOne.Equals(targetEntity) && !IsEntityInSkipList(relation.EntityMany) && !relatedEntities.Contains(relation.EntityMany)) relatedEntities.Add(relation.EntityMany);
                if (relation.EntityMany.Equals(targetEntity) && !IsEntityInSkipList(relation.EntityOne) && !relatedEntities.Contains(relation.EntityOne)) relatedEntities.Add(relation.EntityOne);
            }
            return relatedEntities;
        }

        // Keep adding entity names if entities are used for relationship with the given entities
        private  void ExpandTargetEntities(List<string> targetEntities, int depth)
        {
            if (depth == 0) return;
            List<string> addedEntities = new List<string>();
            foreach (string entity in targetEntities)
            {
                foreach (ERRelation relation in ERRelations.Where(rel => entity.Equals(rel.EntityOne) || entity.Equals(rel.EntityMany)))
                {
                    if (!targetEntities.Contains(relation.EntityOne) && !addedEntities.Contains(relation.EntityOne) && !IsEntityInSkipList(relation.EntityOne)) addedEntities.Add(relation.EntityOne);
                    if (!targetEntities.Contains(relation.EntityMany) && !addedEntities.Contains(relation.EntityMany) && !IsEntityInSkipList(relation.EntityMany)) addedEntities.Add(relation.EntityMany);
                }
            }
            targetEntities.AddRange(addedEntities);
            ExpandTargetEntities(targetEntities, depth - 1);
        }

        internal IEnumerable<ERRelation> GetERRelationsForSpecificEntity(string targetEntity)
        {
            return ERRelations.Where(rel => rel.EntityOne.Equals(targetEntity) || rel.EntityMany.Equals(targetEntity));
        }

        internal IEnumerable<EREntityAttribute> GetEREntitieAttributesByEntityName(string targetEntity)
        {
            return EREntitieAttributes.Where(def => def.EntityName.Equals(targetEntity));
        }

        internal bool IsEntityInSkipList(string entityName)
        {
            // Some entities (such as SystemUser entity) are not useful for ER diagram. Skip those noisy not useful entities for visualization.
            return (SKIP_ENTITIES_FOR_GRAPH.Contains($",{entityName},"));
        }
    }





    internal static class ERInformationUtil
    {
        internal static ERInformation GetERInformationForAllEntities(string organizationUrl, List<EntityMetadata> allEntitiesMetadata)
        {
            // allEntitiesMetadata has all metadata now. 
            // Aggregate all entities' ERInformation and create a final ERInformation object, then persist the data in json.
            ERInformation erInformationForAllEntities = new ERInformation();
            erInformationForAllEntities.OrganizationUrl = organizationUrl;
            erInformationForAllEntities.CreatedOn = DateTime.UtcNow;
            erInformationForAllEntities.DataSourcePath = "API";
            foreach (EntityMetadata entityMetadata in allEntitiesMetadata)
            {
                erInformationForAllEntities.AddMetadata(entityMetadata.ERInformationForThisEntity);
            }
            return erInformationForAllEntities;
        }

        internal static void AggregateERInformationForAllEntities(string organizationUrl, List<EntityMetadata> allEntitiesMetadata, StreamWriter jsonWriter, StreamWriter specWriter, string dataSourcePath = "")
        {
            // allEntitiesMetadata has all metadata now. 
            // Aggregate all entities' ERInformation and create a final ERInformation object, then persist the data in json.
            ERInformation erInformationForAllEntities = new ERInformation();
            erInformationForAllEntities.OrganizationUrl = organizationUrl;
            erInformationForAllEntities.CreatedOn = DateTime.UtcNow;
            erInformationForAllEntities.DataSourcePath = dataSourcePath;
            foreach (EntityMetadata entityMetadata in allEntitiesMetadata)
            {
                erInformationForAllEntities.AddMetadata(entityMetadata.ERInformationForThisEntity);
            }
            string relationsJson = JsonConvert.SerializeObject(erInformationForAllEntities, Formatting.Indented);
            jsonWriter.WriteLine(relationsJson);
            DumpERRelationsInSpecFile(specWriter, erInformationForAllEntities);
        }

        private static void DumpERRelationsInSpecFile(StreamWriter specWriter, ERInformation entityRelations)
        {
            // Add all relationship information in the scheme (spec file) at the end
            foreach (ERRelation relation in entityRelations.ERRelations.Distinct().OrderBy(obj => obj.EntityOne))
            {
                specWriter.WriteLine("{0,50} 1-n {1}.{2}", relation.EntityOne, relation.EntityMany, relation.AttributeMany);
            }
        }

        internal static ERInformation LoadRelationsFromDataFile(string fileNamePath)
        {
            string relationJsonText;
            using (StreamReader jsonTextReader = new StreamReader(fileNamePath))
            {
                relationJsonText = jsonTextReader.ReadToEnd();
            }
            ERInformation entityRelations = JsonConvert.DeserializeObject<ERInformation>(relationJsonText);
            return entityRelations;
        }
    }
/*        internal static ERInformation LoadInformationFromSQLServer()
        {
            string txtServer = Globals.ThisAddIn.registryKey.GetValue("ImportServerName", "edi-dwuat-tmp").ToString();
            string txtDatabase = Globals.ThisAddIn.registryKey.GetValue("ImportDBName", "DataVault").ToString();

            Utilites.ScreenEvents.ShowInputDialog(ref txtServer, ref txtDatabase, "", "", "Choose Server and Database to import.");

            string connectionString = "Data Source=" + txtServer + ";Initial Catalog=" = txtDatabase + ";Integrated Security=true";

            using (var connection2 = new SqlConnection(connectionString))
            {
                var enityReturn = connection2.Query<EREntity>("").ToList();
                if(enityReturn == null)
                {
                    System.Windows.Forms.MessageBox.Show("No Data found, are you sure this is the correct database?");
                    return;
                }
                Globals.ThisAddIn.drawingManager.ERI.EREntities = enityReturn;
            }
*/
            /*
            using (SqlConnection connection = new SqlConnection(connectionString))
            {

                string queryString = "Select * from Information_Schema.Tables";
                SqlCommand command = new SqlCommand(queryString, connection);
                connection.Open();
                SqlDataReader sqlDataReader = command.ExecuteReader();

                if (!sqlDataReader.HasRows())
                {
                    System.Windows.Forms.MessageBox.Show("No Data found, are you sure this is the correct database?");
                    return;
                }

                ERInformation entityRelations = new ERInformation();

                DataTable dataTable = new DataTable("Tables");
                dataTable.Load(sqlDataReader);
                dataTable.Columns.Add("MetadataID", typeof(String));
                dataSet.Tables.Add(dataTable);



                queryString = "select col.*, IsPrimary = iif(cu.column_name is null, 0, 1) from INFORMATION_SCHEMA.columns col left join INFORMATION_SCHEMA.KEY_COLUMN_USAGE cu on " +
                    " col.table_schema = cu.table_schema and col.table_name = cu.table_name and col.column_name = cu.column_name and " +
                    "   OBJECTPROPERTY(OBJECT_ID(cu.CONSTRAINT_SCHEMA + '.' + QUOTENAME(cu.CONSTRAINT_NAME)), 'IsPrimaryKey') = 1 AND cu.TABLE_NAME = 'TableName' AND cu.TABLE_SCHEMA = 'Schema'";
                command = new SqlCommand(queryString, connection);
                //               connection.Open();
                sqlDataReader = command.ExecuteReader();

                dataTable = new DataTable("Columns");
                dataTable.Load(sqlDataReader);
                dataTable.Columns.Add("MetadataID", typeof(String));
                dataSet.Tables.Add(dataTable);



                queryString = "select schema_name(fk_tab.schema_id) + '.' + fk_tab.name as foreign_table," +
                    "    schema_name(pk_tab.schema_id) + '.' + pk_tab.name as primary_table," +
                    "    substring(column_names, 1, len(column_names) - 1) as [fk_columns]," +
                    "    fk.name as fk_constraint_name" +
                    " from sys.foreign_keys fk" +
                    "    inner join sys.tables fk_tab" +
                    "       on fk_tab.object_id = fk.parent_object_id" +
                    "   inner join sys.tables pk_tab " +
                    "       on pk_tab.object_id = fk.referenced_object_id" +
                    " cross apply(select col.[name] + ', '" +
                    "          from sys.foreign_key_columns fk_c" +
                    "              inner join sys.columns col" +
                    "                  on fk_c.parent_object_id = col.object_id" +
                    "                  and fk_c.parent_column_id = col.column_id" +
                    "          where fk_c.parent_object_id = fk_tab.object_id" +
                    "            and fk_c.constraint_object_id = fk.object_id" +
                    "                  order by col.column_id" +
                    "                            for xml path ('') ) D(column_names)" +
                    " order by schema_name(fk_tab.schema_id) + '.' + fk_tab.name," +
                    "    schema_name(pk_tab.schema_id) + '.' + pk_tab.name";

                command = new SqlCommand(queryString, connection);
                //              connection.Open();
                sqlDataReader = command.ExecuteReader();

                dataTable = new DataTable("Relations");
                dataTable.Load(sqlDataReader);
                dataTable.Columns.Add("MetadataID", typeof(String));
                dataSet.Tables.Add(dataTable);
            }*/
        
    

}
