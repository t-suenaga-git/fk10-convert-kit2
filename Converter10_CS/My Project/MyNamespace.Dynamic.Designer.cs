using System;
using System.ComponentModel;
using System.Diagnostics;
using Converter10.Njc.Frm;

namespace Converter10.My
{
    internal static partial class MyProject
    {
        internal partial class MyForms
        {

            [EditorBrowsable(EditorBrowsableState.Never)]
            public MainFrm m_MainFrm;

            public MainFrm MainFrm
            {
                [DebuggerHidden]
                get
                {
                    m_MainFrm = Create__Instance__<MainFrm>(m_MainFrm);
                    return m_MainFrm;
                }
                [DebuggerHidden]
                set
                {
                    if (object.ReferenceEquals(value, m_MainFrm))
                        return;
                    if (value is not null)
                        throw new ArgumentException("Property can only be set to Nothing");
                    Dispose__Instance__<MainFrm>(ref m_MainFrm);
                }
            }

        }


    }
}