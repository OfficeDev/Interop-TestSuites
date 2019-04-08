namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Reflection;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// This class is used to convert specified sub-request type to generic sub-request type and convert generic sub-response type to specified sub-response type.
    /// </summary>
    public static class FsshttpConverter
    {
        /// <summary>
        /// Convert SubResponseElementGenericType to special sub response.
        /// </summary>
        /// <typeparam name="T">The type of special sub response.</typeparam>
        /// <param name="value">The instance of SubResponseElementGenericType.</param>
        /// <returns>The instance of special sub response.</returns>
        public static T ConvertToSpecialSubResponse<T>(object value)
        {
            T targetObject = Activator.CreateInstance<T>();

            PropertyInfo[] props = targetObject.GetType().GetProperties();

            foreach (PropertyInfo item in props)
            {
                object propValue = GetSpecifiedPropertyValueByName(value, item.Name);
                if (string.Compare("SubResponseData", item.Name, StringComparison.OrdinalIgnoreCase) == 0 && propValue != null)
                {
                    object subResponseDataValue = null;
                    switch (item.PropertyType.Name)
                    {
                        case "SchemaLockSubResponseDataType":
                            subResponseDataValue = ConvertToSpecialSubResponse<SchemaLockSubResponseDataType>(propValue);
                            break;
                        case "CellSubResponseDataType":
                            subResponseDataValue = ConvertToSpecialSubResponse<CellSubResponseDataType>(propValue);
                            break;
                        case "CoauthSubResponseDataType":
                            subResponseDataValue = ConvertToSpecialSubResponse<CoauthSubResponseDataType>(propValue);
                            break;
                        case "ExclusiveLockSubResponseDataType":
                            subResponseDataValue = ConvertToSpecialSubResponse<ExclusiveLockSubResponseDataType>(propValue);
                            break;
                        case "ServerTimeSubResponseDataType":
                            subResponseDataValue = ConvertToSpecialSubResponse<ServerTimeSubResponseDataType>(propValue);
                            break;
                        case "WhoAmISubResponseDataType":
                            subResponseDataValue = ConvertToSpecialSubResponse<WhoAmISubResponseDataType>(propValue);
                            break;
                        case "GetDocMetaInfoSubResponseDataType":
                            subResponseDataValue = ConvertToSpecialSubResponse<GetDocMetaInfoSubResponseDataType>(propValue);
                            break;
                        case "EditorsTableSubResponseTypeSubResponseData":
                            subResponseDataValue = ConvertToSpecialSubResponse<EditorsTableSubResponseTypeSubResponseData>(propValue);
                            break;
                        case "VersioningSubResponseDataType":
                            subResponseDataValue = ConvertToSpecialSubResponse<VersioningSubResponseDataType>(propValue);
                            break;
                        case "LockStatusSubResponseDataType":
                            subResponseDataValue = ConvertToSpecialSubResponse<LockStatusSubResponseDataType>(propValue);
                            break;
                        case "Object":
                            subResponseDataValue = null;
                            break;
                        default:
                            throw new InvalidOperationException("Invalid SubResponseData type " + item.PropertyType.Name);
                    }

                    SetSpecifiedProtyValueByName(targetObject, item.Name, subResponseDataValue);
                }
                else
                {
                    SetSpecifiedProtyValueByName(targetObject, item.Name, propValue);
                }
            }

            return targetObject;
        }

        /// <summary>
        /// Convert special sub request to target object.
        /// </summary>
        /// <typeparam name="T">The type of target object.</typeparam>
        /// <param name="value">The instance of special sub request.</param>
        /// <returns>The instance of target object.</returns>
        public static T ConvertSubRequestToGenericType<T>(object value)
        {
            T targetObject = Activator.CreateInstance<T>();
            PropertyInfo[] props = value.GetType().GetProperties();

            foreach (PropertyInfo item in props)
            {
                object propValue = GetSpecifiedPropertyValueByName(value, item.Name);
                if (string.Compare("SubRequestData", item.Name, StringComparison.OrdinalIgnoreCase) == 0 && propValue != null)
                {
                    SetSpecifiedProtyValueByName(targetObject, item.Name, ConvertSubRequestToGenericType<SubRequestDataGenericType>(propValue));
                }
                else
                {
                    PropertyInfo optionalPropertyInGenericType = GetSpecifiedPropertyByName(targetObject, item.Name + "Specified");
                    PropertyInfo optionalPropertyInSubRequestType = GetSpecifiedPropertyByName(value, item.Name + "Specified");
                    if (optionalPropertyInGenericType != null && optionalPropertyInSubRequestType == null)
                    {
                        SetSpecifiedProtyValueByName(targetObject, item.Name + "Specified", true);
                    }

                    SetSpecifiedProtyValueByName(targetObject, item.Name, propValue);
                }
            }

            if (targetObject.GetType() == typeof(SubRequestElementGenericType))
            {
                // set the type of sub request
                SubRequestAttributeType subRequestType = new SubRequestAttributeType();

                switch (value.GetType().Name)
                {
                    case "CellSubRequestType":
                        subRequestType = SubRequestAttributeType.Cell;
                        break;
                    case "CoauthSubRequestType":
                        subRequestType = SubRequestAttributeType.Coauth;
                        break;
                    case "ExclusiveLockSubRequestType":
                        subRequestType = SubRequestAttributeType.ExclusiveLock;
                        break;
                    case "SchemaLockSubRequestType":
                        subRequestType = SubRequestAttributeType.SchemaLock;
                        break;
                    case "ServerTimeSubRequestType":
                        subRequestType = SubRequestAttributeType.ServerTime;
                        break;
                    case "WhoAmISubRequestType":
                        subRequestType = SubRequestAttributeType.WhoAmI;
                        break;
                    case "EditorsTableSubRequestType":
                        subRequestType = SubRequestAttributeType.EditorsTable;
                        break;
                    case "GetDocMetaInfoSubRequestType":
                        subRequestType = SubRequestAttributeType.GetDocMetaInfo;
                        break;
                    case "GetVersionsSubRequestType":
                        subRequestType = SubRequestAttributeType.GetVersions;
                        break;
                    case "VersioningSubRequestType":
                        subRequestType = SubRequestAttributeType.Versioning;
                        break;
                    case "FileOperationSubRequestType":
                        subRequestType = SubRequestAttributeType.FileOperation;
                        break;
                    case "LockStatusSubRequestType":
                        subRequestType = SubRequestAttributeType.LockStatus;
                        break;
                    default:
                        throw new InvalidOperationException("Invalid object type " + value.GetType().Name);
                }

                SetSpecifiedProtyValueByName(targetObject, "Type", subRequestType);
            }

            return targetObject;
        }

        /// <summary>
        /// Set a value in the target object using the specified property name
        /// </summary>
        /// <param name="target">The target object</param>
        /// <param name="propertyName">The property name</param>
        /// <param name="value">The property value</param>
        public static void SetSpecifiedProtyValueByName(object target, string propertyName, object value)
        {
            if (string.IsNullOrEmpty(propertyName) || null == value || null == target)
            {
                return;
            }

            PropertyInfo matchedProperty = GetSpecifiedPropertyByName(target, propertyName);
            if (matchedProperty != null)
            {
                matchedProperty.SetValue(target, value, null);
            }
            else
            {
                throw new InvalidOperationException("Cannot find the property name in the target type " + target.GetType().Name);
            }
        }

        /// <summary>
        /// Get a value in the target object using the specified property name
        /// </summary>
        /// <param name="target">The target object</param>
        /// <param name="propertyName">The property name value</param>
        /// <returns>The property value</returns>
        public static object GetSpecifiedPropertyValueByName(object target, string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName) || null == target)
            {
                return null;
            }

            PropertyInfo matchedProperty = GetSpecifiedPropertyByName(target, propertyName);
            object value = null;
            if (matchedProperty != null)
            {
                value = matchedProperty.GetValue(target, null);
            }

            return value;
        }

        /// <summary>
        /// Get a value in the target object using the specified property name
        /// </summary>
        /// <param name="target">The target object</param>
        /// <param name="propertyName">The property name value</param>
        /// <returns>The property value</returns>
        public static PropertyInfo GetSpecifiedPropertyByName(object target, string propertyName)
        {
            Type currentType = target.GetType();
            PropertyInfo property = currentType.GetProperty(propertyName);
            return property;
        }
    }
}