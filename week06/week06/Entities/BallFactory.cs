﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using week06.Abstractions;

namespace week06.Entities
{
    class BallFactory : IToyFactory
    {
        public Ball CreateNew()
        {
            return new Ball();
        }
    }
}
