package BCDemo;

public class Character {
	
	private String name, age, DOB, notes, eyeCol, hairCol, race, favMovie, favCol, favBook, supNat, mom, dad, bro, sis, loveInt;
	
	public Character(String charName, String charAge, String charDOB, String charNotes, 
			String charEyeColor, String charHairColor, String charRace, String charFavMovie, String charFavColour, String charFavBook, 
			String charSupNat, String charMother, String charFather, String charBrother, String charSister, String charLoveInt){
		name = charName;
		age = charAge;
		DOB= charDOB;
		notes = charNotes;
		
		eyeCol = charEyeColor;
		hairCol = charHairColor;
		race = charRace;
		
		favMovie = charFavMovie;
		favCol = charFavColour;
		favBook = charFavBook;
		
		supNat = charSupNat;
		
		mom = charMother;
		dad = charFather;
		bro = charBrother;
		sis = charSister;
		loveInt = charLoveInt;
	}
	
	public String getName(){
		return name;
	}
	public String getAge(){
		return age;
	}
	public String getBehav(){
		return DOB;
	}
	public String getBio(){
		return notes;
	}
	public String getEyeCol(){
		return eyeCol;
	}
	public String getHairCol(){
		return hairCol;
	}
	public String getRace(){
		return race;
	}
	public String getFavMovie(){
		return favMovie;
	}
	public String getFavCol(){
		return favCol;
	}
	public String getFavBook(){
		return favBook;
	}
	public String getSuper(){
		return supNat;
	}
	public String getMom(){
		return mom;
	}
	public String getDad(){
		return dad;
	}
	public String getBro(){
		return bro;
	}
	public String getSis(){
		return sis;
	}
	public String getLove(){
		return loveInt;
	}
	
	
	
	public void updateCharacter(String charName, String charAge, String charDOB, String charNotes, 
			String charEyeColor, String charHairColor, String charRace, String charFavMovie, String charFavColour, String charFavBook, 
			String charSupNat, String charMother, String charFather, String charBrother, String charSister, String charLoveInt){
		name = charName;
		age = charAge;
		DOB= charDOB;
		notes = charNotes;
		
		eyeCol = charEyeColor;
		hairCol = charHairColor;//
		race = charRace;
		
		favMovie = charFavMovie;
		favCol = charFavColour;
		favBook = charFavBook;
		
		supNat = charSupNat;
		
		mom = charMother;
		dad = charFather;
		bro = charBrother;
		sis = charSister;
		loveInt = charLoveInt;
	}
	
}
