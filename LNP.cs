 /*
Copyright (c) 2017 icebahamut (icebahamut[at]hotmail[dot]com)

Permission is hereby granted, free of charge, to any person obtaining a copy of this
software and associated documentation files (the "Software"), to deal in the Software
without restriction, including without limitation the rights to use, copy, modify, merge,
publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or
substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
DEALINGS IN THE SOFTWARE.
*/

/*
 * Interface Library for Corsair LNP device
 * 
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading;
using HidLibrary;

namespace LNPcmd
{
	/// <summary>
	/// Lighting Node Pro Command 
	/// </summary>
	public class LNP:IDisposable{
		public static bool CheckReplyConfirm = false;
		public static int VENDOR_ID = 0x1B1C;
		public static int[] PRODUCT_ID = {0x0C0B};
		string version = "";
		string fwversion = "";
		string productname = "";
		string sernum = "";
		
		HidDevice device;
		
		LNP(HidDevice device){
			this.device = device;
			this.device.MonitorDeviceEvents = false;
			GetDescription();
			GetLNPDeviceVer();
			GetLNPFWVer();
			GetSerial();
		}
		
		public void Dispose(){
			Close();
		}
		
		public void Close(){
			if(device!=null)device.CloseDevice();
		}
		
		public string GetDevicePath(){
			if(device==null){
				return "";
			}
			return device.DevicePath;
		}
		
		/// <summary>
		/// Get USB HID name
		/// </summary>
		public string GetDescription(){
			if(string.IsNullOrEmpty(productname)){
				if(device==null){
					return productname;
				}
				byte[] b = new byte[0];
				try{
					if(device.ReadProduct(out b)){
						productname = Encoding.Unicode.GetString(b).Replace("\0","");
					}else{
						productname = device.Description;
					}
				}catch(Exception){
					productname = device.Description;
				}
			}
			return productname;
		}
		
		/// <summary>
		/// Get LNP Device Version
		/// </summary>
		public string GetLNPDeviceVer(){
			
			if(device==null){
				return null;
			}
			
			if(string.IsNullOrEmpty(version)){
				byte[] b = new byte[2];
				b[1] = 0x2;
				device.Write(b,500);
				
				var report = device.ReadReport();
				if(report.ReadStatus == HidDeviceData.ReadStatus.Success){
					b = report.Data;
					version = string.Format("{0}.{1}.{2}",b[1],b[2],b[3]);
				}else{
					version = null;
				}
			}
			return version;
		}
		
		/// <summary>
		/// Get LNP Device Firmware
		/// </summary>
		public string GetLNPFWVer(){
			if(device==null){
				return null;
			}
			
			if(string.IsNullOrEmpty(fwversion)){
				byte[] b = new byte[2];
				b[1] = 0x6;
				device.Write(b,500);
				
				var report = device.ReadReport();
				if(report.ReadStatus == HidDeviceData.ReadStatus.Success){
					b = report.Data;
					fwversion = string.Format("{0}.{1}.{2}",b[1],b[2],b[3]);
				}else{
					fwversion = null;
				}
			}
			return fwversion;
		}
		
		/// <summary>
		/// Get Serial Number
		/// </summary>
		/// <returns></returns>
		public string GetSerial(){
			if(string.IsNullOrEmpty(sernum)){
				if(device==null){
					return sernum;
				}
				byte[] b = new byte[0];
				try{
					if(device.ReadSerialNumber(out b)){
						sernum = Encoding.Unicode.GetString(b).Replace("\0","");
					}
				}catch(Exception){
					sernum = "";
				}
			}
			return sernum;
		}
		
		/// <summary>
		/// Send Commit LED change to LNP device, also use for ACK LNP device
		/// </summary>
		/// <returns>Returns true if device received ACK packet</returns>
		public bool SendACK(){
			//if(device==null)return false;
			byte[] b = new byte[3];
			b[1] = 0x33;
			b[2] = 0xFF;
			
			device.Write(b,500);
			if(CheckReplyConfirm){
				b = device.ReadReport().Data;
				if(b[0] == 0){
					return true;
				}
				return false;
			}
			return true;
			
		}
		
		
		public bool SetLEDColorRange(int channel, int index, Color[] color){
			if(device==null)return false;
			
			
				for(int i=0;i<(color.Length/50)+((color.Length%50)==0?0:1);i++){
					if((1+i)*50>=color.Length){
						if(!SetLEDColor(channel, index+(i*50), (color.Length-(index+(i*50))), color)) return false;
					}else{
						if(!SetLEDColor(channel, index+(i*50), 50, color))return false;
					}
				
				}

			return true;
		}
		
		
		
		bool SetLEDColor(int channel, int ledindex, int ledcount, Color[] color){
			if(device==null)return false;
			
			
			if(ledindex<0 || ledindex>128){
				throw new IndexOutOfRangeException("Index is not in range between 0 to 128");
			}
			if(ledcount<0){
				throw new IndexOutOfRangeException("LED Count is less than 0");
			}
			
			if(ledindex+ledcount>128){
				
				throw new IndexOutOfRangeException("Index+Count is more than 128");
			}
			
			bool status = false;
			
			
			byte[] b = new byte[6+ledcount];
			b[1] = 0x32; //set led color
			b[2] = (byte)channel; //channel
			b[3] = (byte)ledindex; //starting LED index
			b[4] = (byte)ledcount; //LED count filling

			for(int i=0;i<3;i++){
				b[5] = (byte)i; //0=R, 1=G, 2=B;
				
				for(int j=0;j<ledcount;j++){
					b[6+j] = (byte)((color[ledindex+j].ToArgb()>>(8*(2-i)))&0xFF); // value
				}
				device.Write(b,500);
				if(CheckReplyConfirm){
					byte[] ret = device.Read(500).Data;
					if(ret[0] == 0){
						status = true;
					}else{
						return false;
					}
				}
			}
			
			if(!CheckReplyConfirm)return true;
			return status;
		}
		
		
		
		
		
		/// <summary>
		/// Reset LNP Channel to zero state(all LED off)
		/// </summary>
		/// <param name="channel">LNP's channel</param>
		/// <returns>Return true if LNP accepts changes</returns>
		public bool ResetChannel(int channel){
			byte[] b = new byte[3];
			b[1] = 0x37;//reset LED to empty
			b[2] = (byte)channel;
			device.Write(b,500);
			
			if(CheckReplyConfirm){
				b = device.Read(500).Data;
				if(b[0] == 0){
					return true;
				}
				return false;
			}
			return true;
		}
		
		/// <summary>
		/// Config Demo Gradient speed and direction on specific channel.
		/// When LNP receives no ACK after 1 minute or more, the LNP will reset itself and enter demo using this config
		/// </summary>
		/// <param name="channel">LNP's channel</param>
		/// <param name="ledcount">LED count. WARNING: setting more than 128 are seriously danger and will cause LNP lockup</param>
		/// <param name="speed">demo animation speed 0-3, higher slower</param>
		/// <param name="reverse">demo in reverse frame</param>
		/// <returns>Return true if LNP accepts changes</returns>
		public bool ConfigDemoMode(int channel, int ledcount, int speed, bool reverse){
			if(ledcount>128)ledcount = 128;//limit it so it doesn't cause serious problem
			if(speed<0)speed = 0;
			if(speed>3)speed = 3;//no effect when higher than 3;
			byte[] b = new byte[8];
			b[1] = 0x35;//set demo mode when idling(no ACK), required before 0x37 and after ACK
			b[2] = (byte)channel;
			b[4] = (byte)ledcount;//how many led to demo, anything higher than 0x80(>128) will be bad
			b[6] = (byte)speed;//animation speed, higher value slower
			b[7] = (byte)(reverse?1:0);//0 or 1 = animation direction(anti/clockwise)
			device.Write(b,500);
			if(CheckReplyConfirm){
				b = device.Read(500).Data;
				if(b[0] == 0){
					return true;
				}
				return false;
			
			}
			return true;
		}
		
		/// <summary>
		/// Enable demo mode or disable demo mode on specific channel
		/// </summary>
		/// <param name="channel">LNP's channel</param>
		/// <param name="enable">True enable demo, false disable demo</param>
		/// <returns>Return true if LNP accepts changes</returns>
		public bool SetDemoOnOff(int channel, bool enable){
			byte[] b = new byte[4];
			b[1] = 0x39;//set idle mode
			b[2] = (byte)channel;
			b[3] = (byte)(enable?0x64:0);//0 = turn off led demo, 0x1-0x64 = turn on led demo
			device.Write(b,500);
			
			if(CheckReplyConfirm){
				b = device.Read(500).Data;
				if(b[0] == 0){
					return true;
				}
				return false;
			}
			return true;
		}
		
		
		/// <summary>
		/// Enter Manual Mode for allow using 0x32 command(direct LED color command)
		/// This mode will exit and return to demo mode after 1 minute of idle without any ACK
		/// </summary>
		/// <param name="channel">LNP's channel</param>
		/// <returns>Return true if LNP accepts changes</returns>
		public bool EnterManualMode(int channel){
			byte[] b = new byte[4];
			b[1] = 0x38;
			b[2] = (byte)channel;
			b[3] = 0x2;
			device.Write(b,500);
			
			if(CheckReplyConfirm){
				b = device.Read(500).Data;
				if(b[0] == 0){
					return true;
				}
				return false;
			}
			return true;
		}
		
		/// <summary>
		/// Enter Demo Gradient Mode using ConfigDemoMode setting
		/// In this mode no 0x32 will be accept
		/// </summary>
		/// <param name="channel">LNP's channel</param>
		/// <returns>Return true if LNP accepts changes</returns>
		public bool EnterDemoMode(int channel){
			byte[] b = new byte[4];
			b[1] = 0x38;
			b[2] = (byte)channel;
			b[3] = 0x1;
			device.Write(b,500);
			if(CheckReplyConfirm){
				b = device.Read(500).Data;
				if(b[0] == 0){
					return true;
				}
				return false;
			}
			return true;
		}
		
		
		/// not sure 0x34 
		bool Send_0x34_Command(int channel){
			byte[] b = new byte[4];
			b[1] = 0x34;//?
			b[2] = (byte)channel;
			b[3] = 0x0;
			device.Write(b,500);
			if(CheckReplyConfirm){
				b = device.Read(500).Data;
				if(b[0] == 0){
					return true;
				}
				return false;
			}
			return true;
		}
		
		
		/// <summary>
		/// Open connection to LNP device with USB device path
		/// </summary>
		/// <param name="devicepath">LNP USB device path</param>
		/// <returns>Return connected LNP device</returns>
		public static LNP Connect(string devicepath){
			var device = HidDevices.GetDevice(devicepath);
			if(device.Attributes.VendorId == VENDOR_ID){
				bool sameproduct = false;
				foreach(int id in PRODUCT_ID){
					if(device.Attributes.ProductId == id){
						sameproduct = true;
						break;
					}
				}
				
				if(sameproduct){
					device.OpenDevice(DeviceMode.Overlapped, DeviceMode.Overlapped, ShareMode.Exclusive);
					int count = 5;
					while(!device.IsConnected && !device.IsOpen){
						
						Thread.Sleep(1000);
						count--;
						if(count<0){
							device.CloseDevice();
							throw new Exception("Failed to connect LNP device!");
						}
					}
					
					return new LNP(device);
				}
			}
			throw new Exception("No such LNP device exist or not LNP device!");
		}
		
		
		/// <summary>
		/// List of available LNP device
		/// </summary>
		/// <returns></returns>
		public static Dictionary<string,string> GetListOfLNP(){
			var e = HidDevices.Enumerate(VENDOR_ID,PRODUCT_ID).GetEnumerator();
			var devices = new Dictionary<string,string>();
			while(e.MoveNext()){
				string productname = "";
				byte[] b = new byte[0];
				try{
					if(e.Current.ReadProduct(out b)){
						productname = Encoding.Unicode.GetString(b).Replace("\0","");
					}else{
						productname = e.Current.Description;
					}
				}catch(Exception){
					productname = e.Current.Description;
				}
				devices.Add(e.Current.DevicePath,productname);
			}
			return devices;
		}
	}
}
