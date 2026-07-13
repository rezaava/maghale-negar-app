using System;
namespace MaghaleNegar.Interfaces
{
    internal interface ITransitionCommand
    {
        Action TransitionMovePreviousCommand { get; set; }
        Action TransitionMoveNextCommand { get; set; }

    }
}
