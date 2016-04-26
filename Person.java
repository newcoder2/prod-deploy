package storage;

public class Person {
	
	/*
	 * 
	OVERLOADING:
	You can have same name for a method but different # and type of parameters.
	Constructor overloading: create two constructors within the class, different parameters only.


	OVERRIDING: -> does not apply for constructors
	Two classes, B extends A: All methods from A will be inherited by B.
	If you have same method in both classes, the B class (child) will override the class A method.

	 
	
	Constructors */
	
	public Person(){  
		//nothing to print
	}

	public Person(int x){   
		System.out.print(x);
	}
	
	//ENCAPSULATION================================================================
	private String Name;
	private int Age;
	
	/*Hide information from external users (firstCode Class)
	 * these values can be retrieve and modified using Get / Set methods.
	 */
	public void setName(String Name){
		this.Name=Name; // Name is passed in parameters by calling method setName
	}
	
	public String getName(){
		return Name;
	}
	//=============================================================================
	
	
	//POLYMORPHISM =================================================================
	/* OVERLOADING: We can have same name for methods but with different # and
	types of parameters
	Constructors can be overloaded but not overrided
	*/
	
	public void Move(){
	
		System.out.println("It's only walking");
	}
	
	public void Move(int i){
		System.out.println("has walked " + i + " mts");
	}
	
	public void Move(int i, int y){
		System.out.print(" has walked from km " + i + " to Km "+ y); 
	}
	//=============================================================================
}
