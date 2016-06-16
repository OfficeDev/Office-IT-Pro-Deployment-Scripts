#region

using System;
using System.Reflection;
using System.Web.Mvc;

#endregion

namespace SelfService.Utils
{
    public class Helper
    {
        /// <summary>
        ///     To Copy changed values from one object to another using FormCollection values.
        /// </summary>
        internal static void CopyUpdatedValues(Object currentUser, Object user, FormCollection values)
        {
            foreach (var value in values)
            {
                PropertyInfo sourceProp = user.GetType()
                    .GetProperty(value.ToString(), BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic);
                PropertyInfo destProp = currentUser.GetType()
                    .GetProperty(value.ToString(), BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic);
                if (null != sourceProp && sourceProp.CanRead)
                {
                    if (null != destProp && destProp.CanWrite)
                    {
                        destProp.SetValue(currentUser, sourceProp.GetValue(user, null), null);
                    }
                }
            }
        }
    }
}