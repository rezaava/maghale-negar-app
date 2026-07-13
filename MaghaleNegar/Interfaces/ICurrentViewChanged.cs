using System;
namespace MaghaleNegar.Interfaces
{
    interface ICurrentViewChanged
    {
        Action CurrentViewChanged { get; set; }
    }
}
