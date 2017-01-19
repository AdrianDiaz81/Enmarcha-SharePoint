using System;
using System.Collections.Generic;
using System.Linq;
using Enmarcha.SharePoint.Abstract.Interfaces.Artefacts;
using Microsoft.SharePoint.Client;
using ListTemplateType = Enmarcha.SharePoint.Abstract.Enum.ListTemplateType;
using RoleType = Enmarcha.SharePoint.Abstract.Enum.RoleType;

namespace Enmarcha.SharePoint.Online.Entities.Artefacts
{
    public class ListSharePoint : IListSharePoint
    {
        #region Properties
        public string Name { get; set; }
        public ClientContext Context { get; set; }
        public ILog Logger { get; set; }
        #endregion

        #region Constructor

        public ListSharePoint(ClientContext context, ILog logger)
        {
            Context = context;
            Logger = logger;
        }

        public ListSharePoint(ClientContext context, ILog logger, string name) : this(context, logger)
        {
            Name = name;
        }

        #endregion

        #region Interface
        public bool Create(string description, ListTemplateType type, bool versionControl)
        {
            //TODO versionControl in List
            var result = false;
            try
            {
                var web = Context.Web;
                var listCreationInfo = new ListCreationInformation
                {
                    Title = Name,
                    Description = description,
                    TemplateType = (int) type                    
                };
                web.Lists.Add(listCreationInfo);
                Context.ExecuteQuery();
                result= true;
            }
            catch (Exception exception)
            {
                Logger.Error($"Error Create List:{exception.Message}");
            }
            return result;
        }

        public bool Delete()
        {
            var result = false;
            try
            {
                var web = Context.Web;
                var list = web.Lists.GetByTitle(Name);
                list.DeleteObject();
                Context.ExecuteQuery();
                result = true;
            }
            catch (Exception exception)
            {
                Logger.Error($"Error Delete List:{exception.Message}");
            }
            return result;
        }

        public bool Exist()
        {
            var result = false;
            try
            {
                var web = Context.Web;
                var list = web.Lists.GetByTitle(Name);                
                Context.Load(list);
                Context.ExecuteQuery();
                result = true;
            }
            catch (Exception exception)
            {
                Logger.Error($"Error Delete List:{exception.Message}");
            }
            return result;
        }

        public bool AddContentType(string contentTypeName)
        {
            //TODO Funcion Get Id of ContentType
            var result = false;
            try
            {
                var list = Context.Web.Lists.GetByTitle(Name);                
                var contentType = Context.Web.ContentTypes.GetById(contentTypeName);                
                list.ContentTypes.AddExistingContentType(contentType);
                Context.ExecuteQuery();
                result = true;
            }
            catch (Exception exception)
            {
                Logger.Error($"Error Add ContentType in List:{exception.Message}");
            }
            return result;
        }

        public bool DeleteContentType(string contentTypeName)
        {
            var result = false;
            try
            {
                var list = Context.Web.Lists.GetByTitle(Name);
                var contentType = list.ContentTypes.GetById(contentTypeName);
                contentType.DeleteObject();                
                Context.ExecuteQuery();
                result = true;
            }
            catch (Exception exception)
            {
                Logger.Error($"Error Add ContentType in List:{exception.Message}");
            }
            return result;
        }

        public bool AddPermissionsGroup(string @group, RoleType role)
        {
            throw new System.NotImplementedException();
        }

        public bool RemovePermissionsGroup(string @group)
        {
            throw new System.NotImplementedException();
        }

        public bool ClearPermisions()
        {
            throw new System.NotImplementedException();
        }

        public bool CreateFolder(string name)
        {
            throw new System.NotImplementedException();
        }

        public IEnumerable<string> GetContentType()
        {
            var result = new List<string>();
            try
            {                
                var list = Context.Web.Lists.GetByTitle(Name);
                var contentTypeCollection = list.ContentTypes;
                Context.Load(contentTypeCollection);
                Context.ExecuteQuery();
                result.AddRange(contentTypeCollection.Select(contentType => contentType.Name));
            }
            catch (Exception exception)
            {
                Logger.Error($"Error Add ContentType in List:{exception.Message}");
            }
            return result;
        }

        #endregion
    }
}