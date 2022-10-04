using System;

namespace Rubberduck.Core.WebApi.Model
{
    /// <summary>
    /// Represents the base properties of a data model entity.
    /// </summary>
    public interface IEntity
    {
        /// <summary>
        /// The internal primary key.
        /// </summary>
        int Id { get; set; }
        /// <summary>
        /// The timestamp when the entity was first created.
        /// </summary>
        DateTime DateInserted { get; set; }
        /// <summary>
        /// The timestamp when the entity was last updated.
        /// </summary>
        DateTime? DateUpdated { get; set; }
    }
}
