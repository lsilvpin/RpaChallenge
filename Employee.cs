namespace RpaChallenge
{
  class Employee
  {
    public string firstName { get; set; }
    public string lastName { get; set; }
    public string companyName { get; set; }
    public string roleInCompany { get; set; }
    public string address { get; set; }
    public string email { get; set; }
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
