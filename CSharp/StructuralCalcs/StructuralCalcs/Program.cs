/*
 * Created by SharpDevelop.
 * User: Conrad
 * Date: 6/10/2018
 * Time: 11:51
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;

namespace StructuralCalcs
{
	class Program
	{
		public static void Main(string[] args)
		{
			double Cpe;
			double qz;
			double pn;
			double w;
			double L;
			double M;
						
			Cpe=-0.7;
			qz=0.96;				//kPa
			pn=Cpe*qz;				//kPa
			w=pn*3.000;  			//kN/m
			L=6.000;				//m
			M=w*Math.Pow(L,2)/8;	//kNm

			Console.WriteLine("Moment " + M.ToString("0.00").PadLeft(8) + " kNm"  );
			
		}
	}
}