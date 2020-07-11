using System.ComponentModel;

namespace RpaChallenge
{
  class Employee
  {
    [DisplayNameAttribute("First Name")]
    public string firstName{ get; set; }
    [DisplayNameAttribute("Last Name ")]
    public string lastName { get; set; }
    [DisplayNameAttribute("Company Name")]
    public string companyName { get; set; }
    [DisplayNameAttribute("Role in Company")]
    public string roleInCompany { get; set; }
    [DisplayNameAttribute("Address")]
    public string address { get; set; }
    [DisplayNameAttribute("Email")]
    public string email { get; set; }
    [DisplayNameAttribute("Phone Number")]
    public string phoneNumber { get; set; }

    public Employee() : base()
    {
    }
    public Employee(string firstName, string lastName, string companyName, 
      string roleInCompany, string address, string email, string phoneNumber)
    {
      this.firstName = firstName;
      this.lastName = lastName;
      this.companyName = companyName;
      this.roleInCompany = roleInCompany;
      this.address = address;
      this.email = email;
      this.phoneNumber = phoneNumber;
    }
  }
}
