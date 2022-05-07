import discord
from discord.ext import commands
from core.classes import Cog_Extension

class others(Cog_Extension):
    @commands.command()
    async def clear(self,ctx,number=100):
        number += 1
        await ctx.channel.purge(limit=number)
    @clear.error
    async def clear_error(self,ctx,error):
        if isinstance(error,commands.errors.ConversionError):
            await ctx.send('資料傳換錯誤')
        if isinstance(error,commands.errors.MissingRequiredArgument):
            await ctx.send('所需指令參數遺失,\n使用方式:```$clear [count(int)]```')
        if isinstance(error,commands.errors.ArgumentParsingError):
            await ctx.send('無法判斷輸入指令,\n使用方式:```$clear [count(int)]```')
        if isinstance(error,commands.errors.DisabledCommand):
            await ctx.send('指令已禁用')
        if isinstance(error,commands.errors.MessageNotFound): 
            await ctx.send('訊息已遭到刪除')
        if isinstance(error,commands.errors.BadArgument): 
            await ctx.send('```-輸入指令格式錯誤\n$clear [count(int)]```')
        if isinstance(error,commands.errors.ExpectedClosingQuoteError): 
            await ctx.send('```-輸入指令格式錯誤\n$clear [count(int)]```')    
        #else: await ctx.send('無法識別的錯誤')
def setup(bot):
    bot.add_cog(others(bot))        