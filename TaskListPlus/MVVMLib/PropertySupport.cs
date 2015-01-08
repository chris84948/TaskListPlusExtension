using System;
using System.Linq.Expressions;
using System.Reflection;

namespace chrisbjohnson.TaskListPlus.MVVMLib
{
    /// <summary>
    /// Property Support Class
    /// </summary>
    public static class PropertySupport
    {
        /// <summary>
        /// Used to extract the property name from the expression
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="propertyExpresssion"></param>
        /// <returns></returns>
        public static String ExtractPropertyName<T>(Expression<Func<T>> propertyExpresssion)
        {
            if (propertyExpresssion == null)
            {
                throw new ArgumentNullException("propertyExpresssion");
            }

            var memberExpression = propertyExpresssion.Body as MemberExpression;
            if (memberExpression == null)
            {
                throw new ArgumentException("The expression is not a member access expression.", "propertyExpresssion");
            }

            var property = memberExpression.Member as PropertyInfo;
            if (property == null)
            {
                throw new ArgumentException("The member access expression does not access a property.", "propertyExpresssion");
            }

            var getMethod = property.GetGetMethod(true);
            if (getMethod.IsStatic)
            {
                throw new ArgumentException("The referenced property is a static property.", "propertyExpresssion");
            }

            return memberExpression.Member.Name;
        }
    }
}
