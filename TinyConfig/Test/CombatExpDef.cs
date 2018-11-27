using System.Collections.Generic;

public class CombatExpDef
{
	public int Level;
	public string Name;
	public ExpClass Exp;
	public int[] Property;
	public MonsterClass[, ] Monster;

	public CombatExpDef()
	{
		Exp = new ExpClass();
		Property = new int[2];
		Monster = new MonsterClass[1, 2];
	    for (int c0 = 0; c0 < 1; c0++)
	    for (int c1 = 0; c1 < 2; c1++)
	        Monster[c0, c1] = new MonsterClass();
    }

	public class ExpClass
	{
		public int Min;
		public int Max;
	}

	public class MonsterClass
	{
		public string Name;
	}

	private static Table.Index _index_ = new Table.Index();
	private static List<CombatExpDef> _values_ = new List<CombatExpDef>();
	public static List<CombatExpDef> Values => _values_;

	public static bool Load(string str)
	{
		Table.Parser parser = new Table.Parser(new Table.Parser.Options {space = ',',title = 0, context = 1});
		return parser.Load(str, new Handler());
	}

	public static CombatExpDef Find(int Level)
	{
		var key = _index_.Get();
		key.Clear();
		key.Add(Level.ToString());
		int idx = _index_.Of(key.ToString());
		if (idx < 0) return null;
		return _values_[idx];
	}

	private class Handler : Table.IHandler
	{
		private CombatExpDef def;
		void Table.IHandler.Start()
		{
			def = new CombatExpDef();
		}
		void Table.IHandler.End()
		{
			_values_.Add(def);
			var key = _index_.Get();
			key.Clear();
			key.Add(def.Level.ToString());
			_index_.Add(key.ToString(), _values_.Count-1);
		}
		bool Table.IHandler.Value(string key, string value)
		{
			switch(key)
			{
				case "Level":
					 return StringUtil.Parse(value, out def.Level);
				case "Name":
					 return StringUtil.Parse(value, out def.Name);
				case "Exp#Min":
					 return StringUtil.Parse(value, out def.Exp.Min);
				case "Exp#Max":
					 return StringUtil.Parse(value, out def.Exp.Max);
				case "Property#1":
			        return StringUtil.Parse(value, out def.Property[0]);
			    case "Property#2":
			        return StringUtil.Parse(value, out def.Property[1]);
			    case "Monster#1#1#Name":
			        return StringUtil.Parse(value, out def.Monster[0, 0].Name);
			    case "Monster#1#2#Name":
			        return StringUtil.Parse(value, out def.Monster[0, 1].Name);
            }
			return false;
		}
	}

}
