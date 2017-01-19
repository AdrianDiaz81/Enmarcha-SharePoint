using System;
using System.Linq;
using Enmarcha.SharePoint.Abstract.Interfaces.Artefacts;
using Microsoft.SharePoint.Client;

namespace Enmarcha.SharePoint.Online.Entities.Artefacts
{
    public class ContentType:IContentType
    {
        #region Properties
        public string Name { get; set; }
        public string GroupName { get; set; }
        public string Parent { get; set; }
        public ClientContext Context { get; set; }
        public ILog Logger { get; set; }
        #endregion

        #region Constructor

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="context"></param>
        /// <param name="logger"></param>
        public ContentType(ClientContext context, ILog logger)
        {
            Context = context;
            Logger = logger;
        }

        public ContentType(ClientContext context, ILog logger, string name, string groupName, string parent) : this(context, logger)
        {
            Name = name;
            GroupName = groupName;
            Parent = parent;
        }
        #endregion

        #region Interface
        public bool Create()
        {
          return  Create(string.Empty);            
        }

        public bool Create(string id)
        {
            try
            {
                var contentTypes = Context.Web.ContentTypes;
                Context.Load(contentTypes);
                Context.ExecuteQuery();
                var parentContentType = contentTypes.GetById(Parent);
                Context.Load(parentContentType);
                Context.ExecuteQuery();

                var contentType = new ContentTypeCreationInformation
                {
                    Name = Name,
                    ParentContentType = parentContentType,
                    Group = GroupName,
                    Id = id
                };
                var addContentType = contentTypes.Add(contentType);
                Context.Load(addContentType);
                Context.ExecuteQuery();
                return true;
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Error Create ContentType:", exception.Message));
                return false;
            }
        }

        public bool Delete()
        {
            var result = false;
            try
            {
                if (Exist())
                {
                    var contentTypes = Context.Web.ContentTypes;
                    Context.Load(contentTypes);
                    Context.ExecuteQuery();
                    foreach (var contentType in contentTypes)
                    {
                        if (contentType.Name.Equals(Name))
                        {
                            contentType.DeleteObject();
                            Context.ExecuteQuery();
                            result = true;
                            break;
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Error Delete ContentType:", exception.Message));
            }
            return result;            
        }

        public bool Exist()
        {
            var result = false;
            try
            {
                var contentTypes = Context.Web.ContentTypes;
                Context.Load(contentTypes);
                Context.ExecuteQuery();
                if (Enumerable.Any(contentTypes, contentType => contentType.Name.Equals(Name)))
                {
                    result = true;
                }
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Error Create ContentType:", exception.Message));
            }            
            return result;            
        }

        public bool AddColumn(string name)
        {
            var result = false;
            try
            {
                var fields = Context.Web.Fields;
                var contentTypes = Context.Web.ContentTypes;
                Context.Load(fields);
                Context.Load(contentTypes);
                var ctType = Enumerable.FirstOrDefault(contentTypes, contentType => contentType.Name.Equals(Name));
                var field = fields.GetByInternalNameOrTitle(name);
                var fieldLink = new FieldLinkCreationInformation {Field = field};
                ctType.FieldLinks.Add(fieldLink);
                ctType.Update(true);
                Context.ExecuteQuery();
                result = true;
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Error AddColumn ContentType:", exception.Message));
            }
            return result;
        }

        public bool RemoveColumn(string name)
        {
            var result = false;
            try
            {
                var fields = Context.Web.Fields;
                var contentTypes = Context.Web.ContentTypes;
                Context.Load(fields);
                Context.Load(contentTypes);
                var field = fields.GetByInternalNameOrTitle(name);
                Context.Load(field);
                Context.ExecuteQuery();
                var ctType = Enumerable.FirstOrDefault(contentTypes, contentType => contentType.Name.Equals(Name));
                var cTypeField = ctType.FieldLinks.GetById(field.Id);
                cTypeField.DeleteObject();
                ctType.Update(true);
                Context.Load(ctType);
                Context.ExecuteQuery();
                result = true;
            }
            catch (Exception exception)
            {
                Logger.Error(string.Concat("Error AddColumn ContentType:", exception.Message));
            }
            return result;
        }
        #endregion
    }
}